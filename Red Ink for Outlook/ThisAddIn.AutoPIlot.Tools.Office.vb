' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Tools.Office.vb
' Purpose:
'   Main AutoPilot Office-tool implementation layer for creating and converting Word,
'   Excel and PowerPoint artifacts and for resolving configured Office design resources.
'
' Architecture / Function:
'   - Resolves active design-set catalogs, design metadata, carriers, guidance and style
'     policies before document generation; bootstrap-classified user/source format carriers
'     suppress implicit repository defaults and are validated before creator execution.
'   - Word creation is OOXML-first: slot-bound DOCX designs and generic no-template Word
'     documents are generated without starting Word. Legacy non-slot carriers remain a
'     bounded compatibility path; live Excel workbook tools are isolated in Office.Interop.
'   - Structured Word generation delegates package mutation to OpenXmlTemplate and visual
'     insertion to OpenXmlVisuals, preserving native styles/numbering and validating output.
'   - PowerPoint/Excel creation and conversion keep their existing dedicated paths while
'     artifact registration, path containment, retry fidelity and logging stay common.
'   - Generic presentation parameters must not silently override a structured design's
'     native document structure.
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
        Public Property IsImplicitDefault As Boolean

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

    Private Class AutoPilotPowerPointLayoutCandidate
        Public Property DesignName As String
        Public Property LayoutName As String
        Public Property SelectionReason As String
        Public Property MasterOrdinal As Integer
        Public Property LayoutOrdinal As Integer
        Public Property SlideWidth As Long
        Public Property SlideHeight As Long
        Public Property TextBindings As JArray = New JArray()
        Public Property PlaceholderDetails As JArray = New JArray()
        Public Property SampleSlides As JArray = New JArray()
    End Class

    Private Shared Function GetSourceFormatCreatorRoutingError(args As System.Collections.Generic.Dictionary(Of System.String, System.Object),
                                                                    applicationName As System.String,
                                                                    context As ToolExecutionContext) As System.String
        If context Is Nothing OrElse context.SequencingState Is Nothing OrElse
           Not context.SequencingState.UserSuppliedSourceFormatAuthority Then
            Return System.String.Empty
        End If

        Dim explicitRepositoryDesign As System.Boolean =
            Not System.String.IsNullOrWhiteSpace(GetArgString(args, "design_name"))
        Dim explicitExternalTemplateCarrier As System.Boolean =
            System.String.Equals(applicationName, "PowerPoint", System.StringComparison.OrdinalIgnoreCase) AndAlso
            Not System.String.IsNullOrWhiteSpace(GetArgString(args, "template_attachment_name"))
        Dim repositoryDefaultsExplicitlyDisabled As System.Boolean =
            Not GetArgBool(args, "use_repository_default_design", True)

        If explicitRepositoryDesign OrElse explicitExternalTemplateCarrier OrElse repositoryDefaultsExplicitlyDisabled Then
            Return System.String.Empty
        End If

        Dim reason As System.String = If(context.SequencingState.UserSuppliedSourceFormatAuthorityReason, System.String.Empty).Trim()
        Dim detail As System.String = If(reason = System.String.Empty, System.String.Empty, " Reason: " & reason)
        Return "The user designated a supplied artifact as the authoritative formatting/layout/structure source. " &
               applicationName & " creation with an implicit repository/default design is therefore blocked. " &
               "Preserve or transform the supplied artifact with the appropriate conversion/native transformation path. " &
               "Only if structure-preserving reuse is not possible and a controlled reconstruction is intended, retry the creator with use_repository_default_design=false." & detail
    End Function

    Private Shared Function BuildDesignExecutionNote(design As AutoPilotDesignResolution) As String
        If design Is Nothing OrElse String.IsNullOrWhiteSpace(design.RequestedName) Then Return ""
        If design.Descriptor Is Nothing Then
            Return $" Requested design '{design.RequestedName}' was not found; neutral professional design was used."
        End If
        If design.ApplicationConfig Is Nothing Then
            Return $" Configured design '{design.Descriptor.Name}' has no {design.ApplicationName} profile; neutral professional design was used."
        End If
        If design.Applied Then
            If design.IsImplicitDefault Then
                Return $" Configured default design used: '{design.Descriptor.Name}' ({design.SourceLabel})."
            End If
            Return $" Configured design used: '{design.Descriptor.Name}' ({design.SourceLabel})."
        End If
        Return $" Configured design '{design.Descriptor.Name}' contained no applicable {design.ApplicationName} settings or usable template; neutral professional design was used."
    End Function

    ''' <summary>
    ''' Resolves a named design from AgentResourcesPath[/Local]\designs (explicit profile or discovered carrier),
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
        Dim typedWordRouteRequested As System.Boolean =
            System.String.Equals(applicationName, "Word", System.StringComparison.OrdinalIgnoreCase) AndAlso
            Not System.String.IsNullOrWhiteSpace(GetArgString(args, "document_type"))
        Dim useRepositoryDefaultDesign As System.Boolean = GetArgBool(args, "use_repository_default_design", True)
        Dim hasExplicitExternalTemplateCarrier As System.Boolean =
            System.String.Equals(applicationName, "PowerPoint", System.StringComparison.OrdinalIgnoreCase) AndAlso
            Not System.String.IsNullOrWhiteSpace(GetArgString(args, "template_attachment_name"))
        Dim sourceFormatAuthorityActive As System.Boolean =
            context IsNot Nothing AndAlso
            context.SequencingState IsNot Nothing AndAlso
            context.SequencingState.UserSuppliedSourceFormatAuthority

        If result.RequestedName = "" AndAlso (Not useRepositoryDefaultDesign OrElse hasExplicitExternalTemplateCarrier OrElse sourceFormatAuthorityActive) Then
            If context IsNot Nothing Then
                Dim reason As System.String
                If hasExplicitExternalTemplateCarrier Then
                    reason = "an explicit user/template attachment carrier is present"
                ElseIf sourceFormatAuthorityActive Then
                    reason = "the bootstrap classified the user-supplied artifact as the authoritative format source"
                Else
                    reason = "use_repository_default_design=false"
                End If
                context.Log("Implicit " & applicationName & " repository design suppressed because " & reason & ".", "diag")
            End If
            Return result
        End If

        If result.RequestedName = "" Then
            If System.String.Equals(applicationName, "Word", System.StringComparison.OrdinalIgnoreCase) Then
                Dim requestedDocumentType As System.String = If(GetArgString(args, "document_type"), System.String.Empty).Trim()
                Dim requestedDocumentLanguage As System.String = If(GetArgString(args, "document_language"), System.String.Empty).Trim()
                Dim requestedOrganization As System.String = If(GetArgString(args, "organization"), System.String.Empty).Trim()
                If requestedDocumentType <> System.String.Empty Then
                    result.Descriptor = SharedLibrary.Agents.DesignRepository.FindBestWordDesign(requestedDocumentType, requestedDocumentLanguage, requestedOrganization)
                    If result.Descriptor IsNot Nothing Then
                        result.RequestedName = result.Descriptor.Id
                        args("design_name") = result.RequestedName
                        If context IsNot Nothing Then context.Log("Resolved Word design by document type first: type='" & requestedDocumentType & "', language='" & requestedDocumentLanguage & "' -> '" & result.Descriptor.Name & "'.")
                        ApDashboardLog("Resolved Word design by document type first: '" & result.Descriptor.Name & "'.", "info")
                    End If
                End If
            End If

            If result.Descriptor Is Nothing AndAlso typedWordRouteRequested Then
                result.TemplateWarning = "No unambiguous configured Word design matched document_type='" & If(GetArgString(args, "document_type"), System.String.Empty) & "' and document_language='" & If(GetArgString(args, "document_language"), System.String.Empty) & "'. A global blank/default design will not be substituted for a requested document type."
                If context IsNot Nothing Then context.Log(result.TemplateWarning)
                Return result
            End If

            If result.Descriptor Is Nothing Then
                result.Descriptor = SharedLibrary.Agents.DesignRepository.FindDefaultDesign(applicationName)
                If result.Descriptor Is Nothing Then Return result
                result.RequestedName = result.Descriptor.Id
                result.IsImplicitDefault = True
                args("design_name") = result.RequestedName
                If context IsNot Nothing Then context.Log("Using configured default " & applicationName & " design '" & result.Descriptor.Name & "'.")
                ApDashboardLog("Using configured default " & applicationName & " design '" & result.Descriptor.Name & "'.", "info")
            End If
        Else
            result.Descriptor = SharedLibrary.Agents.DesignRepository.FindDesign(result.RequestedName)
            If result.Descriptor IsNot Nothing AndAlso System.String.Equals(applicationName, "Word", System.StringComparison.OrdinalIgnoreCase) Then
                Dim requestedDocumentType As System.String = If(GetArgString(args, "document_type"), System.String.Empty).Trim()
                If requestedDocumentType <> System.String.Empty Then
                    Dim selectedWord As Newtonsoft.Json.Linq.JObject = result.Descriptor.GetApplicationConfig("Word")
                    Dim selectedType As System.String = If(If(selectedWord Is Nothing, Nothing, selectedWord.Value(Of System.String)("document_type")), System.String.Empty).Trim()
                    If selectedType <> System.String.Empty AndAlso Not System.String.Equals(selectedType, requestedDocumentType, System.StringComparison.OrdinalIgnoreCase) Then
                        Dim requestedDocumentLanguage As System.String = If(GetArgString(args, "document_language"), System.String.Empty).Trim()
                        Dim requestedOrganization As System.String = If(GetArgString(args, "organization"), System.String.Empty).Trim()
                        Dim typeMatched As SharedLibrary.Agents.DocumentDesignDescriptor = SharedLibrary.Agents.DesignRepository.FindBestWordDesign(requestedDocumentType, requestedDocumentLanguage, requestedOrganization)
                        If typeMatched IsNot Nothing Then
                            If context IsNot Nothing Then context.Log("Corrected conflicting model design selection by document type: '" & result.Descriptor.Name & "' -> '" & typeMatched.Name & "'.")
                            result.Descriptor = typeMatched
                            result.RequestedName = typeMatched.Id
                            args("design_name") = result.RequestedName
                        End If
                    End If
                End If
            End If
        End If
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

            Dim validationIndex As Integer = 0
            For Each slideObj As JObject In slidesArray.OfType(Of JObject)()
                validationIndex += 1
                ExpandAutoPilotPowerPointSlideData(slideObj)
                Dim validationError As String = ValidateAutoPilotPowerPointSlideContent(slideObj, validationIndex)
                If validationError <> "" Then
                    response.Success = False
                    response.Response = validationError
                    Return response
                End If
            Next

            Dim requestedTextHeavy As Boolean = GetArgBool(toolCall.Arguments, "allow_text_heavy", False)
            Dim requestedVisualHeavy As Boolean = GetArgBool(toolCall.Arguments, "allow_visual_heavy", False)
            Dim allowTextHeavy As Boolean = requestedTextHeavy AndAlso IsExplicitPowerPointExtremeStyleRequest(context, "text")
            Dim allowVisualHeavy As Boolean = requestedVisualHeavy AndAlso IsExplicitPowerPointExtremeStyleRequest(context, "visual")
            If context IsNot Nothing AndAlso requestedTextHeavy AndAlso Not allowTextHeavy Then context.Log("Ignored model-supplied allow_text_heavy because the current user request did not explicitly ask for a text-heavy deck.", "diag")
            If context IsNot Nothing AndAlso requestedVisualHeavy AndAlso Not allowVisualHeavy Then context.Log("Ignored model-supplied allow_visual_heavy because the current user request did not explicitly ask for a visual-heavy deck.", "diag")
            NormalizeAutoPilotPowerPointSlidePlan(slidesArray, allowTextHeavy, allowVisualHeavy, context)

            ' Preserve an explicitly requested PowerPoint design/template across retries.
            ' A prior branded attempt must never silently degrade to a neutral retry.
            Dim requestedDesignName As String = GetArgString(toolCall.Arguments, "design_name")
            Dim requestedTemplateAttachmentName As String = GetArgString(toolCall.Arguments, "template_attachment_name")
            If context IsNot Nothing Then
                If String.IsNullOrWhiteSpace(requestedDesignName) AndAlso
                   String.IsNullOrWhiteSpace(requestedTemplateAttachmentName) Then
                    If Not String.IsNullOrWhiteSpace(context.RequiredPowerPointDesignName) Then
                        toolCall.Arguments("design_name") = context.RequiredPowerPointDesignName
                        requestedDesignName = context.RequiredPowerPointDesignName
                        context.Log($"PowerPoint retry preserved required design: {requestedDesignName}.", "diag")
                    ElseIf Not String.IsNullOrWhiteSpace(context.RequiredPowerPointTemplateAttachmentName) Then
                        toolCall.Arguments("template_attachment_name") = context.RequiredPowerPointTemplateAttachmentName
                        requestedTemplateAttachmentName = context.RequiredPowerPointTemplateAttachmentName
                        context.Log($"PowerPoint retry preserved required template: {requestedTemplateAttachmentName}.", "diag")
                    End If
                Else
                    If Not String.IsNullOrWhiteSpace(context.RequiredPowerPointDesignName) AndAlso
                       Not String.IsNullOrWhiteSpace(requestedDesignName) AndAlso
                       Not String.Equals(context.RequiredPowerPointDesignName, requestedDesignName, StringComparison.OrdinalIgnoreCase) Then
                        response.Success = False
                        response.Response = $"PowerPoint retry attempted to replace required design '{context.RequiredPowerPointDesignName}' with '{requestedDesignName}'. The original design remains binding for this run."
                        Return response
                    End If
                    If Not String.IsNullOrWhiteSpace(context.RequiredPowerPointTemplateAttachmentName) AndAlso
                       Not String.IsNullOrWhiteSpace(requestedTemplateAttachmentName) AndAlso
                       Not String.Equals(context.RequiredPowerPointTemplateAttachmentName, requestedTemplateAttachmentName, StringComparison.OrdinalIgnoreCase) Then
                        response.Success = False
                        response.Response = $"PowerPoint retry attempted to replace required template '{context.RequiredPowerPointTemplateAttachmentName}' with '{requestedTemplateAttachmentName}'. The original template remains binding for this run."
                        Return response
                    End If
                End If

                If Not String.IsNullOrWhiteSpace(requestedDesignName) Then context.RequiredPowerPointDesignName = requestedDesignName
                If Not String.IsNullOrWhiteSpace(requestedTemplateAttachmentName) Then context.RequiredPowerPointTemplateAttachmentName = requestedTemplateAttachmentName
            End If

            Dim design As AutoPilotDesignResolution = ResolveAutoPilotDocumentDesign(
                toolCall.Arguments,
                "PowerPoint",
                New String() {"style_preset", "accent_color", "secondary_color", "font_name", "aspect_ratio", "footer_text", "show_slide_numbers", "text_color", "muted_color", "light_color", "line_color", "green_color", "red_color", "amber_color", "preserve_template_slides"},
                New String() {".pptx", ".potx"},
                context)

            Dim sourceFormatRoutingError As System.String = GetSourceFormatCreatorRoutingError(toolCall.Arguments, "PowerPoint", context)
            If Not System.String.IsNullOrWhiteSpace(sourceFormatRoutingError) Then
                response.Success = False
                response.ErrorMessage = sourceFormatRoutingError
                response.Response = sourceFormatRoutingError
                If context IsNot Nothing Then context.Log(sourceFormatRoutingError, "warn")
                Return response
            End If

            If context IsNot Nothing AndAlso context.SequencingState IsNot Nothing Then
                SharedLibrary.Agents.ToolCallSequencing.CaptureRetryInvariantArguments(toolCall.ToolName, toolCall.Arguments, context.SequencingState)
            End If

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

            Dim templateAssignments As New Dictionary(Of Integer, AutoPilotPowerPointLayoutCandidate)()
            If templatePath IsNot Nothing Then
                templateAssignments = Await ResolveAutoPilotPowerPointTemplateAssignmentsAsync(
                    templatePath,
                    slidesArray,
                    0,
                    design,
                    context,
                    ct
                ).ConfigureAwait(False)
                Await EnsureUIThread().ConfigureAwait(False)

                If templateAssignments.Count <> slidesArray.Count Then
                    Dim missingSlides As New List(Of String)()
                    For ordinal As Integer = 1 To slidesArray.Count
                        If Not templateAssignments.ContainsKey(ordinal) Then missingSlides.Add(ordinal.ToString(Globalization.CultureInfo.InvariantCulture))
                    Next
                    response.Success = False
                    response.Response =
                        $"PowerPoint template binding could not safely resolve a native master/layout for generated slide(s) {String.Join(", ", missingSlides)}. " &
                        "The tool will not silently fall back to the template's Blank layout when a template/design was requested."
                    context.Log($"PowerPoint template binding blocked: assigned={templateAssignments.Count}/{slidesArray.Count}; missing={String.Join(",", missingSlides)}")
                    Return response
                End If
            End If

            context.Log($"Creating PowerPoint presentation: {fileName} ({slidesArray.Count} slides)" &
                        If(templatePath IsNot Nothing, $" from template: {templateName}", ""))
            ApDashboardLog($"📊 Creating PowerPoint: {fileName}", "step")

            ' Read design-specific rich-composition parameters once. The renderer itself remains
            ' organization-agnostic; concrete proportions live in the design Markdown.
            Dim richGuidance As String = ReadAutoPilotPowerPointDesignGuidance(design, templatePath, context)
            Dim richSettings As Dictionary(Of String, String) = ParsePowerPointGuidanceVisualSettings(richGuidance)

            Const ppLayoutBlank As Integer = 12
            Const ppSaveAsOpenXMLPresentation As Integer = 24

            Dim liveValidationError As String = ""
            Dim nativeCreationError As String = ""
            Dim nativeCreationStage As String = "initializing PowerPoint"
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
                                                                  nativeCreationStage = "opening PowerPoint template"
                                                                  pres = app.Presentations.Open(templatePath, ReadOnly:=0, Untitled:=-1, WithWindow:=0)
                                                                  If templateFromDesign AndAlso Not GetArgBool(toolCall.Arguments, "preserve_template_slides", False) Then
                                                                      For templateSlideIndex As Integer = CInt(pres.Slides.Count) To 1 Step -1
                                                                          Try : pres.Slides(templateSlideIndex).Delete() : Catch ex As System.Exception : End Try
                                                                      Next
                                                                  End If
                                                              Else
                                                                  nativeCreationStage = "creating blank PowerPoint presentation"
                                                                  pres = app.Presentations.Add(0)
                                                                  ApplyAutoPilotPowerPointPageSetup(pres, toolCall.Arguments)
                                                              End If

                                                              If Not String.IsNullOrWhiteSpace(presTitle) Then
                                                                  Try : pres.BuiltInDocumentProperties("Title").Value = presTitle : Catch : End Try
                                                              End If

                                                              Dim existingSlideCount As Integer = CInt(pres.Slides.Count)
                                                              Dim slideIndex As Integer = existingSlideCount
                                                              For Each slideObj As JObject In slidesArray.OfType(Of JObject)()
                                                                  slideIndex += 1
                                                                  Dim sld As Object = Nothing
                                                                  Try
                                                                      Dim generatedOrdinal As Integer = slideIndex - existingSlideCount
                                                                      nativeCreationStage = $"rendering generated slide {generatedOrdinal}"
                                                                      Dim templateLayout As AutoPilotPowerPointLayoutCandidate = Nothing
                                                                      If templatePath IsNot Nothing AndAlso templateAssignments.ContainsKey(generatedOrdinal) Then
                                                                          templateLayout = templateAssignments(generatedOrdinal)
                                                                      End If

                                                                      Dim templateLayoutApplied As Boolean = False
                                                                      sld = AddAutoPilotPowerPointSlide(
                                                                          pres,
                                                                          slideIndex,
                                                                          ppLayoutBlank,
                                                                          templateLayout,
                                                                          templateLayoutApplied
                                                                      )

                                                                      If templateLayout IsNot Nothing Then
                                                                          context.Log(
                                                                              $"PowerPoint template layout selected: master='{templateLayout.DesignName}'; layout='{templateLayout.LayoutName}'; semantic='{NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), slideIndex, existingSlideCount)}'; reason={templateLayout.SelectionReason}; applied={templateLayoutApplied.ToString().ToLowerInvariant()}"
                                                                          )
                                                                          If Not templateLayoutApplied Then
                                                                              Throw New System.Exception(
                                                                                  $"The selected PowerPoint template layout '{templateLayout.LayoutName}' could not be applied to generated slide {generatedOrdinal}. The presentation was not created with a silent Blank-layout fallback."
                                                                              )
                                                                          End If
                                                                      End If

                                                                      RenderAutoPilotPowerPointSlide(
                                                                          pres,
                                                                          sld,
                                                                          slideObj,
                                                                          slideIndex,
                                                                          existingSlideCount,
                                                                          toolCall.Arguments,
                                                                          templateLayoutApplied,
                                                                          templateLayout,
                                                                          richSettings,
                                                                          context
                                                                      )
                                                                      ApplyPowerPointNotes(sld, slideObj.Value(Of String)("notes"))
                                                                  Finally
                                                                      If sld IsNot Nothing Then
                                                                          Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld) : Catch : End Try
                                                                      End If
                                                                  End Try
                                                              Next

                                                              nativeCreationStage = "saving PowerPoint presentation"
                                                              pres.SaveAs(outputPath, ppSaveAsOpenXMLPresentation)

                                                              ' Validate the live PowerPoint object while it is still owned by this
                                                              ' UI-thread automation session. Starting a second PowerPoint COM server
                                                              ' immediately after SaveAs/Close proved unreliable and is not evidence
                                                              ' that the file PowerPoint just serialized is invalid.
                                                              nativeCreationStage = "validating live PowerPoint presentation"
                                                              liveValidationError = ValidateAutoPilotPowerPointLivePresentation(
                                                                  pres, slidesArray, templateAssignments, context)
                                                              If Not String.IsNullOrWhiteSpace(liveValidationError) Then Return False
                                                              Return True
                                                          Catch ex As System.Exception
                                                              nativeCreationError = $"{ex.GetType().FullName}: {ex.Message}"
                                                              Debug.WriteLine($"CreatePowerPoint error during {nativeCreationStage}: {nativeCreationError}")
                                                              Try
                                                                  context.LogError(
                                                                      "PowerPoint native creation failed.",
                                                                      details:=$"stage={nativeCreationStage}; error={nativeCreationError}{System.Environment.NewLine}{ex.StackTrace}")
                                                              Catch
                                                              End Try
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
                ' PowerPoint itself has already validated the live presentation object on the UI
                ' thread after SaveAs. Open XML remains diagnostic-only: it may reveal package issues
                ' without turning vendor/template extension markup into a false-negative delivery gate.
                DiagnoseAutoPilotPowerPointOpenXmlPackage(outputPath, context)

                RegisterAutoPilotGeneratedOutputFile(outputPath)
                Dim templateNote As String = If(templatePath IsNot Nothing, $", based on template '{templateName}'", "")
                Dim designNote As String = BuildDesignExecutionNote(design)
                response.Success = True
                response.Response = $"PowerPoint presentation created: {fileName} ({slidesArray.Count} new slides{templateNote}, {New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply.{designNote}"
                ApDashboardLog($"✓ PowerPoint created: {fileName}", "info")
            Else
                response.Success = False
                If File.Exists(outputPath) Then
                    Try
                        File.Delete(outputPath)
                    Catch ex As System.Exception
                    End Try
                End If
                If Not String.IsNullOrWhiteSpace(liveValidationError) Then
                    response.Response = liveValidationError
                ElseIf Not String.IsNullOrWhiteSpace(nativeCreationError) Then
                    response.ErrorMessage = $"PowerPoint native creation failed during {nativeCreationStage}: {nativeCreationError}"
                    response.Response = response.ErrorMessage
                Else
                    response.Response = "Failed to create PowerPoint presentation."
                End If
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

        If Not response.Success Then
            Dim recoveryMessage As String = (If(response.ErrorMessage, "") & " " & If(response.Response, "")).Trim()
            response.RepairLoopRecoverable =
                recoveryMessage.StartsWith("PowerPoint visual-quality guard:", StringComparison.OrdinalIgnoreCase) OrElse
                recoveryMessage.IndexOf("PowerPoint rich-content density guard:", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                recoveryMessage.IndexOf("PowerPoint template binding could not safely resolve", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                (recoveryMessage.IndexOf("PowerPoint slide ", StringComparison.OrdinalIgnoreCase) >= 0 AndAlso
                 recoveryMessage.IndexOf("contains no renderable body/structured content", StringComparison.OrdinalIgnoreCase) >= 0)
        End If

        Return response
    End Function

    ''' <summary>
    ''' Expands the optional shallow tool-contract field data_json into the richer
    ''' in-process slide object consumed by the renderer. Keeping nested presentation
    ''' data out of the public tool schema improves interoperability across tool-calling
    ''' model providers while preserving the full renderer feature set.
    ''' Explicit slide properties always win over values supplied through data_json.
    ''' </summary>
    Private Shared Function ValidateAutoPilotPowerPointSlideContent(slideObj As JObject, slideIndex As Integer) As String
        If slideObj Is Nothing Then Return $"PowerPoint slide {slideIndex} is not a valid slide object."

        Dim semanticLayout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
        If semanticLayout = "kpis" Then semanticLayout = "kpi"
        If semanticLayout = "org" OrElse semanticLayout = "organization" Then semanticLayout = "structure"

        Dim body As String = If(slideObj.Value(Of String)("body"), "")
        Dim hasBody As Boolean = Not String.IsNullOrWhiteSpace(body)
        Dim hasContent As Boolean = hasBody

        Select Case semanticLayout
            Case "", "title", "section", "closing"
                Return ""
            Case "bullets"
                hasContent = hasBody
            Case "two_column"
                hasContent =
                    Not String.IsNullOrWhiteSpace(slideObj.Value(Of String)("left_title")) OrElse
                    Not String.IsNullOrWhiteSpace(slideObj.Value(Of String)("left_body")) OrElse
                    Not String.IsNullOrWhiteSpace(slideObj.Value(Of String)("right_title")) OrElse
                    Not String.IsNullOrWhiteSpace(slideObj.Value(Of String)("right_body")) OrElse
                    hasBody
            Case "kpi"
                Dim kpis As JArray = TryCast(slideObj("kpis"), JArray)
                hasContent = hasBody OrElse (kpis IsNot Nothing AndAlso kpis.Count > 0)
            Case "table"
                hasContent = hasBody OrElse TryCast(slideObj("table"), JObject) IsNot Nothing
            Case "chart"
                hasContent = hasBody OrElse TryCast(slideObj("chart"), JObject) IsNot Nothing
            Case "cards"
                Dim cards As JArray = TryCast(slideObj("cards"), JArray)
                hasContent = hasBody OrElse (cards IsNot Nothing AndAlso cards.Count > 0)
            Case "process"
                Dim steps As JArray = TryCast(slideObj("steps"), JArray)
                hasContent = hasBody OrElse (steps IsNot Nothing AndAlso steps.Count > 0)
            Case "structure"
                hasContent = hasBody OrElse TryCast(slideObj("structure"), JObject) IsNot Nothing
            Case "timeline"
                Dim events As JArray = TryCast(slideObj("events"), JArray)
                If events Is Nothing Then events = TryCast(slideObj("timeline"), JArray)
                If events Is Nothing Then
                    Dim timelineObject As JObject = TryCast(slideObj("timeline"), JObject)
                    If timelineObject IsNot Nothing Then events = TryCast(timelineObject("events"), JArray)
                End If
                hasContent = hasBody OrElse (events IsNot Nothing AndAlso events.Count > 0)
            Case "comparison"
                hasContent = hasBody OrElse TryCast(slideObj("comparison"), JObject) IsNot Nothing
            Case "matrix"
                hasContent = hasBody OrElse TryCast(slideObj("matrix"), JObject) IsNot Nothing
            Case "quote"
                hasContent = hasBody OrElse Not String.IsNullOrWhiteSpace(slideObj.Value(Of String)("quote"))
            Case Else
                hasContent = hasBody
        End Select

        If Not hasContent AndAlso HasAutoPilotPowerPointStructuredPayloadInDataJson(slideObj, semanticLayout) Then
            hasContent = True
        End If

        If hasContent Then Return ""
        Return $"PowerPoint slide {slideIndex} ('{If(slideObj.Value(Of String)("title"), "")}') uses layout '{semanticLayout}' but contains no renderable body/structured content. Supply body or the layout's required data_json payload; the tool will not create a silently empty content slide."
    End Function

    Private Shared Function ValidateAutoPilotPowerPointVisualQuality(slidesArray As JArray,
                                                                       allowTextHeavy As Boolean,
                                                                       allowVisualHeavy As Boolean) As String
        If slidesArray Is Nothing OrElse slidesArray.Count < 6 Then Return ""

        Dim contentCount As Integer = 0
        Dim bulletCount As Integer = 0
        Dim textSlideCount As Integer = 0
        Dim visualCount As Integer = 0
        Dim consecutiveBullets As Integer = 0
        Dim maxConsecutiveBullets As Integer = 0
        Dim consecutiveVisuals As Integer = 0
        Dim maxConsecutiveVisuals As Integer = 0

        For Each slideObj As JObject In slidesArray.OfType(Of JObject)()
            Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
            If layout = "kpis" Then layout = "kpi"
            If layout = "org" OrElse layout = "organization" Then layout = "structure"

            If layout = "" OrElse layout = "title" OrElse layout = "section" OrElse layout = "closing" Then
                consecutiveBullets = 0
                consecutiveVisuals = 0
                Continue For
            End If

            contentCount += 1
            Dim isRichVisual As Boolean = False
            Select Case layout
                Case "kpi", "chart", "cards", "process", "structure", "timeline", "comparison", "matrix"
                    isRichVisual = True
            End Select

            If isRichVisual Then
                visualCount += 1
                consecutiveVisuals += 1
                consecutiveBullets = 0
                If consecutiveVisuals > maxConsecutiveVisuals Then maxConsecutiveVisuals = consecutiveVisuals
            Else
                textSlideCount += 1
                consecutiveVisuals = 0
                If layout = "bullets" Then
                    bulletCount += 1
                    consecutiveBullets += 1
                    If consecutiveBullets > maxConsecutiveBullets Then maxConsecutiveBullets = consecutiveBullets
                Else
                    consecutiveBullets = 0
                End If
            End If
        Next

        If contentCount < 5 Then Return ""

        Dim minVisualSlides As Integer = If(allowTextHeavy, 0, Math.Max(2, CInt(Math.Ceiling(contentCount * 0.3R))))
        Dim maxVisualSlides As Integer = If(allowVisualHeavy, contentCount, Math.Max(2, CInt(Math.Floor(contentCount * 0.6R))))
        Dim minTextSlides As Integer = If(allowVisualHeavy, 0, Math.Max(2, contentCount - maxVisualSlides))
        Dim maxBulletSlides As Integer = If(allowTextHeavy, contentCount, Math.Max(3, CInt(Math.Ceiling(contentCount * 0.6R))))
        Dim textBalanceOk As Boolean = allowTextHeavy OrElse (bulletCount <= maxBulletSlides AndAlso maxConsecutiveBullets <= 3)
        Dim visualBalanceOk As Boolean = allowVisualHeavy OrElse (visualCount <= maxVisualSlides AndAlso textSlideCount >= minTextSlides AndAlso maxConsecutiveVisuals <= 2)
        Dim minimumVisualOk As Boolean = allowTextHeavy OrElse visualCount >= minVisualSlides

        If textBalanceOk AndAlso visualBalanceOk AndAlso minimumVisualOk Then Return ""

        Return $"PowerPoint visual-quality guard: rebalance this {slidesArray.Count}-slide deck instead of making every content slide graphical or every slide textual. Current mix: {visualCount} rich visual / {textSlideCount} native text-data slides across {contentCount} content slides; bullets={bulletCount}; longest rich-visual run={maxConsecutiveVisuals}; longest bullet run={maxConsecutiveBullets}. For a normal executive deck, use roughly 30-60% rich visuals and deliberately keep native bullets/two-column/table/quote slides for explanatory, legal, tax, caveat, and detail-heavy content. Use structure/process/comparison/cards/timeline/matrix/chart/KPI only where the visual form genuinely improves the story, and normally place no more than two rich visuals consecutively. Set allow_text_heavy=true or allow_visual_heavy=true only when the user explicitly requested that extreme style."
    End Function

    Private Shared Function BuildAutoPilotPowerPointBulletItemsFromLines(lines As System.Collections.Generic.IEnumerable(Of String)) As JArray
        Dim result As New JArray()
        If lines Is Nothing Then Return result

        Dim indentFollowingNumberedHeading As Boolean = False
        For Each raw As String In lines
            Dim cleaned As String = If(raw, "").Trim()
            cleaned = System.Text.RegularExpressions.Regex.Replace(
                cleaned,
                "^(?:[•·▪◦‣⁃∙●○■□◆◇►▸\-\*\+]\s*)+",
                "").TrimStart()
            If cleaned = "" Then Continue For

            Dim level As Integer = 0
            Dim numberedMatch As System.Text.RegularExpressions.Match =
                System.Text.RegularExpressions.Regex.Match(cleaned, "^\s*\d+[\.\)]\s+(.+)$")
            If numberedMatch.Success Then
                cleaned = numberedMatch.Groups(1).Value.Trim()
                level = 0
                indentFollowingNumberedHeading = True
            ElseIf indentFollowingNumberedHeading Then
                level = 1
            End If

            result.Add(New JObject From {
                {"text", cleaned},
                {"level", level}
            })
        Next
        Return result
    End Function

    Private Shared Function GetAutoPilotPowerPointBulletItems(slideObj As JObject) As JArray
        If slideObj Is Nothing Then Return New JArray()

        Dim explicitItems As JArray = TryCast(slideObj("bullet_items"), JArray)
        If explicitItems IsNot Nothing AndAlso explicitItems.Count > 0 Then
            Dim normalized As New JArray()
            For Each itemToken As JToken In explicitItems
                If itemToken Is Nothing OrElse itemToken.Type = JTokenType.Null Then Continue For
                If itemToken.Type = JTokenType.Object Then
                    Dim itemObj As JObject = DirectCast(itemToken, JObject)
                    Dim itemText As String = CleanPptBulletText(If(itemObj.Value(Of String)("text"), "")).Trim()
                    If itemText = "" Then Continue For
                    Dim level As Integer = itemObj.Value(Of Integer?)("level").GetValueOrDefault(0)
                    level = Math.Max(0, Math.Min(4, level))
                    normalized.Add(New JObject From {{"text", itemText}, {"level", level}})
                Else
                    Dim itemText As String = CleanPptBulletText(itemToken.ToString()).Trim()
                    If itemText <> "" Then normalized.Add(New JObject From {{"text", itemText}, {"level", 0}})
                End If
            Next
            Return normalized
        End If

        Dim legacyPoints As JArray = TryCast(slideObj("bullet_points"), JArray)
        If legacyPoints IsNot Nothing AndAlso legacyPoints.Count > 0 Then
            Dim legacyLines As New System.Collections.Generic.List(Of String)()
            For Each pointToken As JToken In legacyPoints
                If pointToken Is Nothing OrElse pointToken.Type = JTokenType.Null Then Continue For
                legacyLines.Add(pointToken.ToString())
            Next
            Return BuildAutoPilotPowerPointBulletItemsFromLines(legacyLines)
        End If

        Dim body As String = If(slideObj.Value(Of String)("body"), "")
        If String.IsNullOrWhiteSpace(body) Then Return New JArray()
        Return BuildAutoPilotPowerPointBulletItemsFromLines(
            body.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, System.StringSplitOptions.RemoveEmptyEntries))
    End Function

    Private Shared Function BuildAutoPilotPowerPointBulletBody(items As JArray) As String
        If items Is Nothing OrElse items.Count = 0 Then Return ""
        Dim lines As New System.Collections.Generic.List(Of String)()
        For Each itemObj As JObject In items.OfType(Of JObject)()
            Dim itemText As String = If(itemObj.Value(Of String)("text"), "").Trim()
            If itemText <> "" Then lines.Add(itemText)
        Next
        Return String.Join(vbCrLf, lines)
    End Function

    Private Shared Function FlattenAutoPilotPowerPointObjectArray(token As JToken) As JArray
        Dim result As New JArray()
        If token Is Nothing OrElse token.Type = JTokenType.Null Then Return result

        Dim pending As New System.Collections.Generic.Stack(Of JToken)()
        pending.Push(token)
        While pending.Count > 0
            Dim current As JToken = pending.Pop()
            If current Is Nothing OrElse current.Type = JTokenType.Null Then Continue While

            If current.Type = JTokenType.Array Then
                Dim arrayToken As JArray = DirectCast(current, JArray)
                For index As Integer = arrayToken.Count - 1 To 0 Step -1
                    pending.Push(arrayToken(index))
                Next
            ElseIf current.Type = JTokenType.Object Then
                result.Add(current.DeepClone())
            End If
        End While

        Return result
    End Function

    Private Shared Sub NormalizeAutoPilotPowerPointTypedVisualContract(visual As JObject)
        If visual Is Nothing Then Exit Sub

        ' Tool-calling providers occasionally add an extra array layer even when the JSON schema
        ' declares a flat object array. Treat this as a contract-normalization issue rather than
        ' rejecting otherwise complete semantic content and forcing an LLM repair loop.
        For Each propertyName As String In New String() {"items", "nodes", "columns", "quadrants", "series"}
            Dim token As JToken = visual(propertyName)
            If token Is Nothing OrElse token.Type = JTokenType.Null Then Continue For

            Dim flattened As JArray = FlattenAutoPilotPowerPointObjectArray(token)
            If flattened.Count > 0 Then visual(propertyName) = flattened
        Next
    End Sub

    Private Shared Sub ExpandAutoPilotPowerPointTypedVisual(slideObj As JObject)
        If slideObj Is Nothing Then Exit Sub
        Dim visual As JObject = TryCast(slideObj("visual"), JObject)
        If visual Is Nothing Then Exit Sub

        NormalizeAutoPilotPowerPointTypedVisualContract(visual)

        Dim visualType As String = If(visual.Value(Of String)("type"), "").Trim().ToLowerInvariant()
        Select Case visualType
            Case "org_chart", "hierarchy"
                visualType = "structure"
            Case "list"
                visualType = "cards"
        End Select
        If visualType = "" Then Exit Sub

        slideObj("layout") = visualType

        Select Case visualType
            Case "cards"
                Dim sourceItems As JArray = TryCast(visual("items"), JArray)
                If sourceItems IsNot Nothing Then
                    Dim cards As New JArray()
                    For Each itemObj As JObject In sourceItems.OfType(Of JObject)()
                        Dim label As String = If(itemObj.Value(Of String)("title"), itemObj.Value(Of String)("label"))
                        Dim detail As String = If(itemObj.Value(Of String)("body"), itemObj.Value(Of String)("detail"))
                        cards.Add(New JObject From {
                            {"title", If(label, "")},
                            {"body", If(detail, "")},
                            {"badge", If(itemObj.Value(Of String)("badge"), "")},
                            {"tone", If(itemObj.Value(Of String)("tone"), "")}
                        })
                    Next
                    If cards.Count > 0 Then slideObj("cards") = cards
                End If

            Case "process"
                Dim sourceItems As JArray = TryCast(visual("items"), JArray)
                If sourceItems IsNot Nothing Then
                    Dim steps As New JArray()
                    For Each itemObj As JObject In sourceItems.OfType(Of JObject)()
                        Dim label As String = If(itemObj.Value(Of String)("title"), itemObj.Value(Of String)("label"))
                        Dim detail As String = If(itemObj.Value(Of String)("body"), itemObj.Value(Of String)("detail"))
                        steps.Add(New JObject From {{"title", If(label, "")}, {"body", If(detail, "")}})
                    Next
                    If steps.Count > 0 Then slideObj("steps") = steps
                End If

            Case "structure"
                Dim nodes As JArray = TryCast(visual("nodes"), JArray)
                If nodes IsNot Nothing AndAlso nodes.Count > 0 Then
                    slideObj("structure") = New JObject From {{"nodes", nodes.DeepClone()}}
                End If

            Case "timeline"
                Dim sourceItems As JArray = TryCast(visual("items"), JArray)
                If sourceItems IsNot Nothing Then
                    Dim events As New JArray()
                    For Each itemObj As JObject In sourceItems.OfType(Of JObject)()
                        Dim label As String = If(itemObj.Value(Of String)("label"), "")
                        Dim eventTitle As String = If(itemObj.Value(Of String)("title"), label)
                        Dim detail As String = If(itemObj.Value(Of String)("detail"), itemObj.Value(Of String)("body"))
                        events.Add(New JObject From {
                            {"label", label},
                            {"title", If(eventTitle, "")},
                            {"body", If(detail, "")}
                        })
                    Next
                    If events.Count > 0 Then slideObj("events") = events
                End If

            Case "comparison"
                Dim columns As JArray = TryCast(visual("columns"), JArray)
                If columns IsNot Nothing AndAlso columns.Count > 0 Then
                    slideObj("comparison") = New JObject From {{"columns", columns.DeepClone()}}
                End If

            Case "matrix"
                Dim matrixObj As New JObject()
                For Each key As String In New String() {"x_left", "x_right", "y_top", "y_bottom", "quadrants"}
                    If visual(key) IsNot Nothing Then matrixObj(key) = visual(key).DeepClone()
                Next
                If matrixObj.Count > 0 Then slideObj("matrix") = matrixObj

            Case "kpi"
                Dim sourceItems As JArray = TryCast(visual("items"), JArray)
                If sourceItems IsNot Nothing Then
                    Dim kpis As New JArray()
                    For Each itemObj As JObject In sourceItems.OfType(Of JObject)()
                        kpis.Add(New JObject From {
                            {"label", If(itemObj.Value(Of String)("label"), itemObj.Value(Of String)("title"))},
                            {"value", If(itemObj.Value(Of String)("value"), "")},
                            {"detail", If(itemObj.Value(Of String)("detail"), itemObj.Value(Of String)("body"))}
                        })
                    Next
                    If kpis.Count > 0 Then slideObj("kpis") = kpis
                End If

            Case "chart"
                Dim chartObj As New JObject()
                If visual("chart_type") IsNot Nothing Then chartObj("type") = visual("chart_type").DeepClone()
                If visual("categories") IsNot Nothing Then chartObj("categories") = visual("categories").DeepClone()
                If visual("series") IsNot Nothing Then chartObj("series") = visual("series").DeepClone()
                If chartObj.Count > 0 Then slideObj("chart") = chartObj
        End Select
    End Sub

    Private Shared Function SplitAutoPilotPowerPointBulletLines(text As String) As List(Of String)
        Dim result As New List(Of String)()
        If String.IsNullOrWhiteSpace(text) Then Return result
        For Each raw As String In text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(ControlChars.Lf)
            Dim value As String = CleanPptBulletText(raw).Trim()
            If value <> "" Then result.Add(value)
        Next
        Return result
    End Function

    Private Shared Function IsExplicitPowerPointExtremeStyleRequest(context As ToolExecutionContext, mode As String) As Boolean
        If context Is Nothing OrElse String.IsNullOrWhiteSpace(context.LatestUserRequestRaw) Then Return False
        Dim requestText As String = context.LatestUserRequestRaw.Trim().ToLowerInvariant()
        Dim normalizedMode As String = If(mode, "").Trim().ToLowerInvariant()
        If normalizedMode = "text" Then
            Return requestText.Contains("text-heavy") OrElse requestText.Contains("text heavy") OrElse
                   requestText.Contains("textlastig") OrElse requestText.Contains("nur text") OrElse
                   requestText.Contains("mostly text") OrElse requestText.Contains("primarily text")
        End If
        If normalizedMode = "visual" Then
            Return requestText.Contains("visual-heavy") OrElse requestText.Contains("visual heavy") OrElse
                   requestText.Contains("grafiklastig") OrElse requestText.Contains("grafik-lastig") OrElse
                   requestText.Contains("nur grafiken") OrElse requestText.Contains("mostly visual") OrElse
                   requestText.Contains("infographic-heavy")
        End If
        Return False
    End Function

    Private Shared Sub NormalizeAutoPilotPowerPointSlidePlan(slidesArray As JArray,
                                                              allowTextHeavy As Boolean,
                                                              allowVisualHeavy As Boolean,
                                                              context As ToolExecutionContext)
        If slidesArray Is Nothing OrElse slidesArray.Count < 6 Then Exit Sub

        Dim richLayouts As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase) From {
            "kpi", "chart", "cards", "process", "structure", "timeline", "comparison", "matrix"
        }
        Dim contentSlides As New List(Of JObject)()
        Dim slideOrdinal As Integer = 0
        For Each slideObj As JObject In slidesArray.OfType(Of JObject)()
            slideOrdinal += 1
            Dim rawLayout As String = If(slideObj.Value(Of String)("layout"), "").Trim()
            Dim layout As String = NormalizeAutoPilotPowerPointSemanticLayout(rawLayout, slideOrdinal, 0)
            If String.IsNullOrWhiteSpace(rawLayout) Then slideObj("layout") = layout

            ' A semantic rich label is not enough: normalize it to an actually renderable payload.
            ' Otherwise the balance counter can count a visual that later silently renders as bullets.
            If layout = "cards" AndAlso TryCast(slideObj("cards"), JArray) Is Nothing Then
                Dim cardLines As List(Of String) = SplitAutoPilotPowerPointBulletLines(slideObj.Value(Of String)("body"))
                If cardLines.Count >= 2 AndAlso cardLines.Count <= 6 Then
                    Dim cards As New JArray()
                    For Each cardLine As String In cardLines.Take(4)
                        Dim cardTitle As String = cardLine
                        Dim cardBody As String = cardLine
                        Dim colon As Integer = cardLine.IndexOf(":"c)
                        If colon > 0 AndAlso colon < 55 Then
                            cardTitle = cardLine.Substring(0, colon).Trim()
                            cardBody = cardLine.Substring(colon + 1).Trim()
                        ElseIf cardTitle.Length > 42 Then
                            Dim titleLength As Integer = 42
                            Dim lastSpace As Integer = cardTitle.LastIndexOf(" "c, titleLength - 1)
                            If lastSpace > 0 Then titleLength = lastSpace
                            cardTitle = cardTitle.Substring(0, titleLength).Trim()
                        End If
                        cards.Add(New JObject From {{"title", cardTitle}, {"body", cardBody}})
                    Next
                    slideObj("cards") = cards
                Else
                    slideObj("layout") = "bullets"
                    layout = "bullets"
                End If
            ElseIf layout = "process" AndAlso TryCast(slideObj("steps"), JArray) Is Nothing Then
                Dim processLines As List(Of String) = SplitAutoPilotPowerPointBulletLines(slideObj.Value(Of String)("body"))
                If processLines.Count >= 2 AndAlso processLines.Count <= 6 Then
                    Dim steps As New JArray()
                    For Each processLine As String In processLines.Take(5)
                        Dim stepTitle As String = processLine
                        Dim stepBody As String = processLine
                        Dim colon As Integer = processLine.IndexOf(":"c)
                        If colon > 0 AndAlso colon < 55 Then
                            stepTitle = processLine.Substring(0, colon).Trim()
                            stepBody = processLine.Substring(colon + 1).Trim()
                        ElseIf stepTitle.Length > 38 Then
                            Dim titleLength As Integer = 38
                            Dim lastSpace As Integer = stepTitle.LastIndexOf(" "c, titleLength - 1)
                            If lastSpace > 0 Then titleLength = lastSpace
                            stepTitle = stepTitle.Substring(0, titleLength).Trim()
                        End If
                        steps.Add(New JObject From {{"title", stepTitle}, {"body", stepBody}})
                    Next
                    slideObj("steps") = steps
                Else
                    slideObj("layout") = "bullets"
                    layout = "bullets"
                End If
            ElseIf richLayouts.Contains(layout) AndAlso Not HasAutoPilotPowerPointStructuredPayload(slideObj, layout) Then
                slideObj("layout") = "bullets"
                layout = "bullets"
            End If

            If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then contentSlides.Add(slideObj)
        Next
        If contentSlides.Count < 5 Then Exit Sub

        Dim targetMin As Integer = If(allowTextHeavy, 0, Math.Max(2, CInt(Math.Ceiling(contentSlides.Count * 0.3R))))
        Dim targetMax As Integer = If(allowVisualHeavy, contentSlides.Count, Math.Max(2, CInt(Math.Floor(contentSlides.Count * 0.6R))))
        Dim visualCount As Integer = System.Linq.Enumerable.Count(contentSlides, Function(slide As JObject) richLayouts.Contains(If(slide.Value(Of String)("layout"), "")))
        If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer started: content={contentSlides.Count}; rich={visualCount}; target={targetMin}-{targetMax}.", "diag")

        ' If the model over-illustrates the deck, deterministically demote selected rich slides
        ' to native text. Preserve all source text; never silently discard structured content.
        If visualCount > targetMax Then
            For index As Integer = contentSlides.Count - 1 To 0 Step -1
                If visualCount <= targetMax Then Exit For
                Dim slideObj As JObject = contentSlides(index)
                Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
                If Not richLayouts.Contains(layout) Then Continue For
                If index > 0 AndAlso index < contentSlides.Count - 1 AndAlso (index Mod 2 = 0) Then Continue For
                Dim fallbackBody As String = BuildPowerPointFallbackBody(slideObj)
                If String.IsNullOrWhiteSpace(fallbackBody) Then Continue For
                slideObj("layout") = "bullets"
                slideObj("body") = fallbackBody
                visualCount -= 1
                If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer demoted rich slide '{If(slideObj.Value(Of String)("title"), "")}' to native bullets to restore deck balance.", "diag")
            Next
        End If

        ' If the model produces a text desert, promote only slides whose existing wording
        ' naturally maps to a visual form. This is deterministic and content-preserving.
        If visualCount < targetMin AndAlso Not allowTextHeavy Then
            For Each slideObj As JObject In contentSlides
                If visualCount >= targetMin Then Exit For
                Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
                If layout <> "bullets" Then Continue For
                Dim title As String = If(slideObj.Value(Of String)("title"), "").ToLowerInvariant()
                Dim lines As List(Of String) = GetAutoPilotPowerPointPromotionLines(slideObj)
                If lines.Count < 3 OrElse lines.Count > 5 Then Continue For

                If title.Contains("ablauf") OrElse title.Contains("prozess") OrElse title.Contains("roadmap") OrElse title.Contains("schritt") OrElse title.Contains("timeline") OrElse title.Contains("steps") Then
                    Dim steps As New JArray()
                    For i As Integer = 0 To lines.Count - 1
                        Dim stepLine As String = System.Text.RegularExpressions.Regex.Replace(lines(i), "^\s*\d+[\.\)]\s*", "").Trim()
                        Dim stepTitle As String = stepLine
                        Dim stepBody As String = stepLine
                        Dim stepColon As Integer = stepLine.IndexOf(":"c)
                        If stepColon > 0 AndAlso stepColon < 55 Then
                            stepTitle = stepLine.Substring(0, stepColon).Trim()
                            stepBody = stepLine.Substring(stepColon + 1).Trim()
                        ElseIf stepTitle.Length > 38 Then
                            Dim stepTitleLength As Integer = 38
                            Dim stepLastSpace As Integer = stepTitle.LastIndexOf(" "c, stepTitleLength - 1)
                            If stepLastSpace > 0 Then stepTitleLength = stepLastSpace
                            stepTitle = stepTitle.Substring(0, stepTitleLength).Trim()
                        End If
                        steps.Add(New JObject From {{"title", stepTitle}, {"body", stepBody}})
                    Next
                    slideObj("layout") = "process"
                    slideObj("steps") = steps
                    visualCount += 1
                    If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer promoted '{If(slideObj.Value(Of String)("title"), "")}' from bullets to process.", "diag")
                ElseIf (title.Contains("vergleich") OrElse title.Contains("gegenüberstellung") OrElse title.Contains("comparison") OrElse title.Contains("compare")) AndAlso lines.Count >= 2 Then
                    Dim comparisonColumns As New JArray()
                    For comparisonIndex As Integer = 0 To 1
                        Dim comparisonLine As String = lines(comparisonIndex)
                        Dim comparisonTitle As String = $"Option {comparisonIndex + 1}"
                        Dim comparisonBody As String = comparisonLine
                        Dim comparisonColon As Integer = comparisonLine.IndexOf(":"c)
                        If comparisonColon > 0 AndAlso comparisonColon < 65 Then
                            comparisonTitle = comparisonLine.Substring(0, comparisonColon).Trim()
                            comparisonBody = comparisonLine.Substring(comparisonColon + 1).Trim()
                        End If
                        Dim comparisonItems As New JArray()
                        comparisonItems.Add(comparisonBody)
                        comparisonColumns.Add(New JObject From {{"title", comparisonTitle}, {"items", comparisonItems}})
                    Next
                    For remainderIndex As Integer = 2 To lines.Count - 1
                        Dim remainder As String = lines(remainderIndex)
                        Dim remainderColon As Integer = remainder.IndexOf(":"c)
                        Dim remainderLabel As String = If(remainderColon > 0, remainder.Substring(0, remainderColon).Trim().ToLowerInvariant(), "")
                        Dim remainderBody As String = If(remainderColon > 0, remainder.Substring(remainderColon + 1).Trim(), remainder)
                        Dim secondColumn As JObject = TryCast(comparisonColumns(1), JObject)
                        If remainderLabel.Contains("empfehl") OrElse remainderLabel.Contains("fazit") Then
                            secondColumn("verdict") = remainder
                        Else
                            Dim secondItems As JArray = TryCast(secondColumn("items"), JArray)
                            secondItems.Add(remainder)
                        End If
                    Next
                    slideObj("layout") = "comparison"
                    slideObj("comparison") = New JObject From {{"columns", comparisonColumns}}
                    visualCount += 1
                    If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer promoted '{If(slideObj.Value(Of String)("title"), "")}' from bullets to comparison.", "diag")
                ElseIf Not (title.Contains("steuer") OrElse title.Contains("recht") OrElse title.Contains("compliance") OrElse title.Contains("detail") OrElse title.Contains("tax") OrElse title.Contains("legal")) AndAlso
                       (title.Contains("motive") OrElse title.Contains("treiber") OrElse title.Contains("drivers") OrElse title.Contains("vorteil") OrElse title.Contains("benefit") OrElse title.Contains("risik") OrElse title.Contains("risk") OrElse title.Contains("aspekt") OrElse title.Contains("aspect") OrElse title.Contains("hürde") OrElse title.Contains("challenge")) Then
                    Dim cards As New JArray()
                    For Each line As String In lines.Take(4)
                        Dim cardTitle As String = line
                        Dim cardBody As String = ""
                        Dim colon As Integer = line.IndexOf(":"c)
                        If colon > 0 AndAlso colon < 55 Then
                            cardTitle = line.Substring(0, colon).Trim()
                            cardBody = line.Substring(colon + 1).Trim()
                        Else
                            Dim titleLength As Integer = Math.Min(42, line.Length)
                            If titleLength < line.Length Then
                                Dim lastSpace As Integer = line.LastIndexOf(" "c, titleLength - 1)
                                If lastSpace > 0 Then titleLength = lastSpace
                            End If
                            cardTitle = line.Substring(0, titleLength).Trim()
                            cardBody = line
                        End If
                        cards.Add(New JObject From {{"title", cardTitle}, {"body", cardBody}})
                    Next
                    slideObj("layout") = "cards"
                    slideObj("cards") = cards
                    visualCount += 1
                    If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer promoted '{If(slideObj.Value(Of String)("title"), "")}' from bullets to cards.", "diag")
                End If
            Next
        End If

        ' Final agnostic balance pass: when the deck is still below the visual target,
        ' promote concise non-legal/non-tax bullet groups to editable cards. This uses
        ' existing structured bullet hierarchy and never invents organization-specific content.
        If visualCount < targetMin AndAlso Not allowTextHeavy Then
            For Each slideObj As JObject In contentSlides
                If visualCount >= targetMin Then Exit For
                Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()
                If layout <> "bullets" Then Continue For
                Dim title As String = If(slideObj.Value(Of String)("title"), "").ToLowerInvariant()
                If title.Contains("steuer") OrElse title.Contains("recht") OrElse title.Contains("compliance") OrElse title.Contains("tax") OrElse title.Contains("legal") Then Continue For
                Dim lines As List(Of String) = GetAutoPilotPowerPointPromotionLines(slideObj)
                If lines.Count < 3 OrElse lines.Count > 4 Then Continue For

                Dim cards As New JArray()
                For Each line As String In lines
                    Dim cardTitle As String = line
                    Dim cardBody As String = line
                    Dim colon As Integer = line.IndexOf(":"c)
                    If colon > 0 AndAlso colon < 55 Then
                        cardTitle = line.Substring(0, colon).Trim()
                        cardBody = line.Substring(colon + 1).Trim()
                    ElseIf cardTitle.Length > 42 Then
                        Dim titleLength As Integer = 42
                        Dim lastSpace As Integer = cardTitle.LastIndexOf(" "c, titleLength - 1)
                        If lastSpace > 0 Then titleLength = lastSpace
                        cardTitle = cardTitle.Substring(0, titleLength).Trim()
                    End If
                    cards.Add(New JObject From {{"title", cardTitle}, {"body", cardBody}})
                Next
                slideObj("layout") = "cards"
                slideObj("cards") = cards
                visualCount += 1
                If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer promoted '{If(slideObj.Value(Of String)("title"), "")}' from bullets to cards to meet the balanced visual target.", "diag")
            Next
        End If

        If context IsNot Nothing Then context.Log($"PowerPoint plan normalizer completed: rich={visualCount}/{contentSlides.Count}.", "diag")
    End Sub

    Private Shared Function GetAutoPilotPowerPointPromotionLines(slideObj As JObject) As List(Of String)
        Dim lines As List(Of String) = SplitAutoPilotPowerPointBulletLines(If(slideObj?.Value(Of String)("body"), ""))
        If lines.Count >= 3 AndAlso lines.Count <= 5 Then Return lines

        Dim bulletItems As JArray = TryCast(slideObj?("bullet_items"), JArray)
        If bulletItems Is Nothing OrElse bulletItems.Count = 0 Then Return lines

        Dim grouped As New List(Of String)()
        For Each itemObj As JObject In bulletItems.OfType(Of JObject)()
            Dim itemText As String = If(itemObj.Value(Of String)("text"), "").Trim()
            If itemText = "" Then Continue For
            Dim level As Integer = Math.Max(0, itemObj.Value(Of Integer?)("level").GetValueOrDefault(0))
            If level = 0 OrElse grouped.Count = 0 Then
                grouped.Add(itemText)
            Else
                grouped(grouped.Count - 1) = grouped(grouped.Count - 1) & " — " & itemText
            End If
        Next
        Return grouped
    End Function

    Private Shared Function HasAutoPilotPowerPointStructuredPayload(slideObj As JObject,
                                                                      semanticLayout As String) As Boolean
        If slideObj Is Nothing Then Return False
        Select Case If(semanticLayout, "").Trim().ToLowerInvariant()
            Case "kpi"
                Dim items As JArray = TryCast(slideObj("kpis"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "table"
                Return TryCast(slideObj("table"), JObject) IsNot Nothing
            Case "chart"
                Return TryCast(slideObj("chart"), JObject) IsNot Nothing
            Case "cards"
                Dim items As JArray = TryCast(slideObj("cards"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "process"
                Dim items As JArray = TryCast(slideObj("steps"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "structure"
                Return TryCast(slideObj("structure"), JObject) IsNot Nothing
            Case "timeline"
                Dim items As JArray = TryCast(slideObj("events"), JArray)
                If items Is Nothing Then items = TryCast(slideObj("timeline"), JArray)
                If items IsNot Nothing AndAlso items.Count > 0 Then Return True
                Dim obj As JObject = TryCast(slideObj("timeline"), JObject)
                Return obj IsNot Nothing AndAlso TryCast(obj("events"), JArray) IsNot Nothing
            Case "comparison"
                Return TryCast(slideObj("comparison"), JObject) IsNot Nothing
            Case "matrix"
                Return TryCast(slideObj("matrix"), JObject) IsNot Nothing
            Case Else
                Return False
        End Select
    End Function

    Private Shared Function NormalizeAutoPilotPowerPointValidationText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim normalized As String = value.Replace(ChrW(&HA0), " "c)
        normalized = System.Text.RegularExpressions.Regex.Replace(normalized, "\s+", " ")
        Return normalized.Trim().ToLowerInvariant()
    End Function

    Private Shared Sub AddAutoPilotPowerPointValidationFragments(target As List(Of String), value As String)
        If target Is Nothing OrElse String.IsNullOrWhiteSpace(value) Then Return
        Dim cleaned As String = CleanPptBulletText(value)
        For Each rawLine As String In cleaned.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(New Char() {ChrW(10)}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim line As String = System.Text.RegularExpressions.Regex.Replace(rawLine.Trim(), "^[\s•·▪◦\-–—]+", "").Trim()
            If line.Length < 2 Then Continue For
            If line.Length > 160 Then line = line.Substring(0, 160)
            target.Add(line)
        Next
    End Sub

    Private Shared Function GetAutoPilotPowerPointValidationFragments(slideObj As JObject,
                                                                        slideIndex As Integer) As List(Of String)
        Dim fragments As New List(Of String)()
        If slideObj Is Nothing Then Return fragments

        AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("title"))
        Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), slideIndex, 0)

        Select Case semanticLayout
            Case "title", "closing"
                Dim secondLine As String = If(slideObj.Value(Of String)("subtitle"), "")
                If secondLine = "" Then secondLine = If(slideObj.Value(Of String)("body"), "")
                AddAutoPilotPowerPointValidationFragments(fragments, secondLine)
            Case "section"
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("section_number"))
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("subtitle"))
            Case "two_column"
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("left_title"))
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("left_body"))
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("right_title"))
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("right_body"))
            Case "quote"
                Dim quoteText As String = If(slideObj.Value(Of String)("quote"), "")
                If quoteText = "" Then quoteText = If(slideObj.Value(Of String)("body"), "")
                AddAutoPilotPowerPointValidationFragments(fragments, quoteText)
                AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("attribution"))
            Case "cards"
                Dim validationCards As JArray = TryCast(slideObj("cards"), JArray)
                If validationCards IsNot Nothing Then
                    For Each validationCard As JObject In validationCards.OfType(Of JObject)()
                        AddAutoPilotPowerPointValidationFragments(fragments, validationCard.Value(Of String)("title"))
                        AddAutoPilotPowerPointValidationFragments(fragments, validationCard.Value(Of String)("body"))
                        AddAutoPilotPowerPointValidationFragments(fragments, validationCard.Value(Of String)("badge"))
                    Next
                End If
            Case "structure"
                Dim validationStructure As JObject = TryCast(slideObj("structure"), JObject)
                If validationStructure IsNot Nothing Then
                    Dim validationTop As JObject = TryCast(validationStructure("top"), JObject)
                    If validationTop IsNot Nothing Then
                        AddAutoPilotPowerPointValidationFragments(fragments, validationTop.Value(Of String)("title"))
                        AddAutoPilotPowerPointValidationFragments(fragments, validationTop.Value(Of String)("body"))
                    End If
                    Dim validationChildren As JArray = TryCast(validationStructure("children"), JArray)
                    If validationChildren IsNot Nothing Then
                        For Each validationChild As JObject In validationChildren.OfType(Of JObject)()
                            AddAutoPilotPowerPointValidationFragments(fragments, validationChild.Value(Of String)("title"))
                            AddAutoPilotPowerPointValidationFragments(fragments, validationChild.Value(Of String)("body"))
                        Next
                    End If
                End If
            Case "process"
                Dim validationSteps As JArray = TryCast(slideObj("steps"), JArray)
                If validationSteps IsNot Nothing Then
                    For Each validationStep As JObject In validationSteps.OfType(Of JObject)()
                        AddAutoPilotPowerPointValidationFragments(fragments, validationStep.Value(Of String)("title"))
                        AddAutoPilotPowerPointValidationFragments(fragments, validationStep.Value(Of String)("body"))
                    Next
                End If
            Case "comparison"
                Dim validationComparison As JObject = TryCast(slideObj("comparison"), JObject)
                Dim validationColumns As JArray = TryCast(validationComparison?("columns"), JArray)
                If validationColumns IsNot Nothing Then
                    For Each validationColumn As JObject In validationColumns.OfType(Of JObject)()
                        AddAutoPilotPowerPointValidationFragments(fragments, validationColumn.Value(Of String)("title"))
                        AddAutoPilotPowerPointValidationFragments(fragments, validationColumn.Value(Of String)("verdict"))
                        Dim validationItems As JArray = TryCast(validationColumn("items"), JArray)
                        If validationItems IsNot Nothing Then
                            For Each validationItem As JToken In validationItems
                                AddAutoPilotPowerPointValidationFragments(fragments, validationItem.ToString())
                            Next
                        End If
                    Next
                End If
            Case Else
                If Not HasAutoPilotPowerPointStructuredPayload(slideObj, semanticLayout) Then
                    AddAutoPilotPowerPointValidationFragments(fragments, slideObj.Value(Of String)("body"))
                End If
        End Select

        Return fragments.Distinct(StringComparer.OrdinalIgnoreCase).ToList()
    End Function

    Private Shared Function ValidateAutoPilotPowerPointOutputContent(outputPath As String,
                                                                      slidesArray As JArray,
                                                                      templateAssignments As IDictionary(Of Integer, AutoPilotPowerPointLayoutCandidate),
                                                                      ByRef validationError As String) As Boolean
        validationError = ""
        If String.IsNullOrWhiteSpace(outputPath) OrElse
           slidesArray Is Nothing OrElse
           slidesArray.Count = 0 OrElse
           Not System.IO.File.Exists(outputPath) Then

            validationError = "PowerPoint output validation could not inspect the generated presentation."
            Return False
        End If

        Try
            Using document As DocumentFormat.OpenXml.Packaging.PresentationDocument =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, False)

                Dim presentationPart As DocumentFormat.OpenXml.Packaging.PresentationPart = document.PresentationPart
                Dim slideIdList As DocumentFormat.OpenXml.Presentation.SlideIdList = presentationPart?.Presentation?.SlideIdList
                If presentationPart Is Nothing OrElse slideIdList Is Nothing Then
                    validationError = "PowerPoint output validation found no readable slide list."
                    Return False
                End If

                Dim slideIds As List(Of DocumentFormat.OpenXml.Presentation.SlideId) =
                    slideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)().ToList()
                If slideIds.Count < slidesArray.Count Then
                    validationError = $"PowerPoint output validation expected {slidesArray.Count} generated slides but found only {slideIds.Count} total slides."
                    Return False
                End If

                Dim firstGeneratedIndex As Integer = slideIds.Count - slidesArray.Count
                For ordinal As Integer = 1 To slidesArray.Count
                    Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
                    If slideObj Is Nothing Then Continue For
                    Dim slideId As DocumentFormat.OpenXml.Presentation.SlideId = slideIds(firstGeneratedIndex + ordinal - 1)
                    If slideId.RelationshipId Is Nothing Then
                        validationError = $"PowerPoint output validation could not resolve generated slide {ordinal}."
                        Return False
                    End If

                    Dim slidePart As DocumentFormat.OpenXml.Packaging.SlidePart =
                        TryCast(presentationPart.GetPartById(slideId.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                    If slidePart?.Slide Is Nothing Then
                        validationError = $"PowerPoint output validation could not read generated slide {ordinal}."
                        Return False
                    End If

                    If templateAssignments IsNot Nothing AndAlso templateAssignments.ContainsKey(ordinal) Then
                        Dim expectedLayout As AutoPilotPowerPointLayoutCandidate = templateAssignments(ordinal)
                        Dim actualLayoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = slidePart.SlideLayoutPart
                        If actualLayoutPart Is Nothing Then
                            validationError = $"PowerPoint output validation failed on generated slide {ordinal}: the selected template layout is missing."
                            Return False
                        End If

                        Dim actualLayoutName As String = GetOpenXmlPowerPointLayoutName(actualLayoutPart, expectedLayout.LayoutOrdinal)
                        If Not String.IsNullOrWhiteSpace(expectedLayout.LayoutName) AndAlso
                           Not System.Text.RegularExpressions.Regex.IsMatch(expectedLayout.LayoutName.Trim(), "^Layout\s+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) AndAlso
                           NormalizePowerPointLayoutKey(actualLayoutName) <> NormalizePowerPointLayoutKey(expectedLayout.LayoutName) Then

                            validationError =
                                $"PowerPoint output validation failed on generated slide {ordinal}: expected template layout '{expectedLayout.LayoutName}' but the saved slide uses '{actualLayoutName}'. The presentation was not registered for delivery."
                            Return False
                        End If
                    End If

                    Dim actualText As String = String.Join(
                        " ",
                        slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().Select(Function(t) t.Text)
                    )
                    Dim normalizedActual As String = NormalizeAutoPilotPowerPointValidationText(actualText)

                    Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), ordinal, 0)
                    If semanticLayout = "cards" OrElse semanticLayout = "comparison" OrElse semanticLayout = "structure" OrElse semanticLayout = "process" OrElse semanticLayout = "timeline" OrElse semanticLayout = "matrix" OrElse semanticLayout = "chart" OrElse semanticLayout = "kpi" Then
                        Dim slideCx As Int64 = CLng(presentationPart.Presentation.SlideSize.Cx.Value)
                        Dim slideCy As Int64 = CLng(presentationPart.Presentation.SlideSize.Cy.Value)
                        Dim protectedRects As List(Of AutoPilotPowerPointVisualRect) = GetOpenXmlPowerPointProtectedRects(slidePart)
                        For Each generatedShape As DocumentFormat.OpenXml.Presentation.Shape In slidePart.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
                            Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties = generatedShape.NonVisualShapeProperties?.NonVisualDrawingProperties
                            Dim generatedName As String = If(nv?.Name?.Value, "")
                            If Not generatedName.StartsWith("Red Ink ", StringComparison.OrdinalIgnoreCase) Then Continue For
                            If generatedShape.TextBody IsNot Nothing AndAlso Not generatedShape.TextBody.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().Any() Then
                                validationError = $"PowerPoint output validation failed on generated slide {ordinal}: shape '{generatedName}' contains an invalid empty text body. The presentation was not registered for delivery."
                                Return False
                            End If
                            Dim generatedRect As AutoPilotPowerPointVisualRect
                            If TryGetOpenXmlPowerPointShapeRect(generatedShape, generatedRect) Then
                                If generatedRect.X < 0 OrElse generatedRect.Y < 0 OrElse generatedRect.X + generatedRect.W > slideCx OrElse generatedRect.Y + generatedRect.H > slideCy Then
                                    validationError = $"PowerPoint output validation failed on generated slide {ordinal}: shape '{generatedName}' extends outside the slide bounds."
                                    Return False
                                End If
                                For Each protectedRect As AutoPilotPowerPointVisualRect In protectedRects
                                    If PowerPointVisualRectsOverlap(generatedRect, protectedRect) Then
                                        validationError = $"PowerPoint output validation failed on generated slide {ordinal}: rich shape '{generatedName}' overlaps a protected title/header/footer zone."
                                        Return False
                                    End If
                                Next
                            End If
                        Next
                    End If

                    Dim fragments As List(Of String) = GetAutoPilotPowerPointValidationFragments(slideObj, ordinal)
                    For Each fragment As String In fragments
                        Dim expected As String = NormalizeAutoPilotPowerPointValidationText(fragment)
                        If expected = "" Then Continue For
                        If Not normalizedActual.Contains(expected) Then
                            validationError = $"PowerPoint output validation failed on generated slide {ordinal}: expected visible text was not found ('{fragment}'). The presentation was not registered for delivery."
                            Return False
                        End If
                    Next
                Next
            End Using
        Catch ex As System.Exception
            validationError = $"PowerPoint output validation failed: {ex.Message}"
            Return False
        End Try

        Return True
    End Function

    Private Shared Function NormalizeOpenXmlPowerPointHex(value As String, fallback As String) As String
        Dim raw As String = If(value, "").Trim().TrimStart("#"c)
        If System.Text.RegularExpressions.Regex.IsMatch(raw, "^[0-9A-Fa-f]{6}$") Then Return raw.ToUpperInvariant()
        Return fallback.Trim().TrimStart("#"c).ToUpperInvariant()
    End Function

    Private Shared Function GetNextOpenXmlPowerPointShapeId(shapeTree As DocumentFormat.OpenXml.Presentation.ShapeTree) As UInt32
        Dim maxId As UInt32 = 1UI
        If shapeTree Is Nothing Then Return 2UI
        For Each nv As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties In shapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nv.Id IsNot Nothing AndAlso nv.Id.Value > maxId Then maxId = nv.Id.Value
        Next
        Return maxId + 1UI
    End Function

    Private Shared Function CreateOpenXmlPowerPointTextShape(shapeId As UInt32,
                                                              shapeName As String,
                                                              x As Int64,
                                                              y As Int64,
                                                              cx As Int64,
                                                              cy As Int64,
                                                              fillHex As String,
                                                              lineHex As String,
                                                              textHex As String,
                                                              paragraphs As IEnumerable(Of Tuple(Of String, Integer, Boolean))) As DocumentFormat.OpenXml.Presentation.Shape
        Dim shape As New DocumentFormat.OpenXml.Presentation.Shape()
        shape.NonVisualShapeProperties = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = shapeId, .Name = shapeName},
            New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(New DocumentFormat.OpenXml.Drawing.ShapeLocks() With {.NoGrouping = True}),
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim transform As New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset() With {.X = x, .Y = y},
            New DocumentFormat.OpenXml.Drawing.Extents() With {.Cx = cx, .Cy = cy})
        Dim geometry As New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {
            .Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle
        }
        Dim properties As New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        properties.Append(transform)
        properties.Append(geometry)
        properties.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = fillHex}))
        Dim outline As New DocumentFormat.OpenXml.Drawing.Outline(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = lineHex}))
        properties.Append(outline)
        shape.ShapeProperties = properties

        Dim textBody As New DocumentFormat.OpenXml.Presentation.TextBody()
        textBody.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties() With {
            .Wrap = DocumentFormat.OpenXml.Drawing.TextWrappingValues.Square,
            .LeftInset = 91440,
            .RightInset = 91440,
            .TopInset = 68580,
            .BottomInset = 68580
        })
        textBody.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())
        Dim paragraphAdded As Boolean = False
        For Each paragraphSpec As Tuple(Of String, Integer, Boolean) In paragraphs
            If paragraphSpec Is Nothing OrElse String.IsNullOrWhiteSpace(paragraphSpec.Item1) Then Continue For
            Dim paragraph As New DocumentFormat.OpenXml.Drawing.Paragraph()
            Dim run As New DocumentFormat.OpenXml.Drawing.Run()
            run.RunProperties = New DocumentFormat.OpenXml.Drawing.RunProperties() With {
                .Language = "de-CH",
                .FontSize = paragraphSpec.Item2 * 100,
                .Bold = paragraphSpec.Item3
            }
            run.RunProperties.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = textHex}))
            run.Text = New DocumentFormat.OpenXml.Drawing.Text(paragraphSpec.Item1)
            paragraph.Append(run)
            textBody.Append(paragraph)
            paragraphAdded = True
        Next
        ' p:txBody requires at least one a:p. Connector/decorative shapes carry no text body at all.
        If paragraphAdded Then shape.TextBody = textBody
        Return shape
    End Function

    Private Structure AutoPilotPowerPointVisualRect
        Public X As Int64
        Public Y As Int64
        Public W As Int64
        Public H As Int64
    End Structure

    Private Shared Function ParsePowerPointGuidanceVisualSettings(guidance As String) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(guidance) Then Return result
        For Each rawLine As String In guidance.Replace(vbCrLf, vbLf).Split(ControlChars.Lf)
            Dim line As String = rawLine.Trim()
            If Not line.StartsWith("|", StringComparison.Ordinal) Then Continue For
            Dim cells As String() = line.Trim("|"c).Split("|"c).Select(Function(v) v.Trim().Trim("`"c)).ToArray()
            If cells.Length < 2 Then Continue For
            Dim key As String = cells(0).Trim().ToLowerInvariant()
            If Not key.StartsWith("rich.", StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim value As String = cells(1).Trim()
            If key <> "" AndAlso value <> "" Then result(key) = value
        Next
        Return result
    End Function

    Private Shared Function GetPowerPointGuidanceSettingDouble(settings As Dictionary(Of String, String), key As String, defaultValue As Double) As Double
        If settings Is Nothing OrElse Not settings.ContainsKey(key) Then Return defaultValue
        Dim raw As String = settings(key).Replace("%", "").Trim()
        Dim parsed As Double
        If Double.TryParse(raw, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, parsed) Then Return parsed
        If Double.TryParse(raw, parsed) Then Return parsed
        Return defaultValue
    End Function

    Private Shared Function GetPowerPointGuidanceSettingInteger(settings As Dictionary(Of String, String), key As String, defaultValue As Integer) As Integer
        Return CInt(Math.Round(GetPowerPointGuidanceSettingDouble(settings, key, defaultValue)))
    End Function

    Private Shared Function TryGetOpenXmlPowerPointShapeRect(shape As DocumentFormat.OpenXml.Presentation.Shape,
                                                              ByRef rect As AutoPilotPowerPointVisualRect) As Boolean
        If shape Is Nothing OrElse shape.ShapeProperties Is Nothing OrElse shape.ShapeProperties.Transform2D Is Nothing Then Return False
        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = shape.ShapeProperties.Transform2D
        If xfrm.Offset Is Nothing OrElse xfrm.Extents Is Nothing Then Return False
        rect.X = CLng(xfrm.Offset.X.Value)
        rect.Y = CLng(xfrm.Offset.Y.Value)
        rect.W = CLng(xfrm.Extents.Cx.Value)
        rect.H = CLng(xfrm.Extents.Cy.Value)
        Return rect.W > 0 AndAlso rect.H > 0
    End Function

    Private Shared Function GetOpenXmlPowerPointPlaceholderKind(shape As DocumentFormat.OpenXml.Presentation.Shape) As String
        If shape Is Nothing OrElse shape.NonVisualShapeProperties Is Nothing OrElse shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties Is Nothing Then Return ""
        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)()
        If ph Is Nothing Then Return ""

        Dim rawType As String = ph.GetAttribute("type", "").Value
        If Not String.IsNullOrWhiteSpace(rawType) Then Return rawType.Trim()
        If ph.Type IsNot Nothing Then Return ph.Type.Value.ToString()

        ' Open XML defaults an omitted placeholder type to an object/content placeholder.
        Return "obj"
    End Function

    Private Shared Function HasOpenXmlPowerPointPlaceholder(shape As DocumentFormat.OpenXml.Presentation.Shape) As Boolean
        If shape Is Nothing OrElse shape.NonVisualShapeProperties Is Nothing OrElse shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties Is Nothing Then Return False
        Return shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)() IsNot Nothing
    End Function

    Private Shared Function IsOpenXmlPowerPointProtectedPlaceholder(kind As String) As Boolean
        Dim normalized As String = If(kind, "").Trim().ToLowerInvariant()
        Select Case normalized
            Case "title", "ctrtitle", "subtitle", "ftr", "footer", "dt", "date", "sldnum", "slidenumber", "hdr", "header"
                Return True
        End Select
        Return False
    End Function

    Private Shared Function GetOpenXmlPowerPointPlaceholderIndex(shape As DocumentFormat.OpenXml.Presentation.Shape) As UInt32?
        If shape Is Nothing OrElse shape.NonVisualShapeProperties Is Nothing OrElse shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties Is Nothing Then Return Nothing
        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            shape.NonVisualShapeProperties.ApplicationNonVisualDrawingProperties.GetFirstChild(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)()
        If ph Is Nothing OrElse ph.Index Is Nothing Then Return Nothing
        Return ph.Index.Value
    End Function

    Private Shared Function TryResolveOpenXmlPowerPointPlaceholderRect(slidePart As DocumentFormat.OpenXml.Packaging.SlidePart,
                                                                       layoutShape As DocumentFormat.OpenXml.Presentation.Shape,
                                                                       ByRef rect As AutoPilotPowerPointVisualRect) As Boolean
        If TryGetOpenXmlPowerPointShapeRect(layoutShape, rect) Then Return True
        Dim idx As UInt32? = GetOpenXmlPowerPointPlaceholderIndex(layoutShape)
        If Not idx.HasValue Then Return False
        Dim masterPart As DocumentFormat.OpenXml.Packaging.SlideMasterPart = slidePart?.SlideLayoutPart?.SlideMasterPart
        Dim masterTree As DocumentFormat.OpenXml.Presentation.ShapeTree = masterPart?.SlideMaster?.CommonSlideData?.ShapeTree
        If masterTree Is Nothing Then Return False
        For Each masterShape As DocumentFormat.OpenXml.Presentation.Shape In masterTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim masterIdx As UInt32? = GetOpenXmlPowerPointPlaceholderIndex(masterShape)
            If masterIdx.HasValue AndAlso masterIdx.Value = idx.Value AndAlso TryGetOpenXmlPowerPointShapeRect(masterShape, rect) Then Return True
        Next
        Return False
    End Function

    Private Shared Function GetOpenXmlPowerPointProtectedRects(slidePart As DocumentFormat.OpenXml.Packaging.SlidePart) As List(Of AutoPilotPowerPointVisualRect)
        Dim result As New List(Of AutoPilotPowerPointVisualRect)()
        Dim layoutTree As DocumentFormat.OpenXml.Presentation.ShapeTree = slidePart?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree
        If layoutTree Is Nothing Then Return result
        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In layoutTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            If Not HasOpenXmlPowerPointPlaceholder(shape) Then Continue For
            Dim kind As String = GetOpenXmlPowerPointPlaceholderKind(shape)
            If Not IsOpenXmlPowerPointProtectedPlaceholder(kind) Then Continue For
            Dim rect As AutoPilotPowerPointVisualRect
            If TryResolveOpenXmlPowerPointPlaceholderRect(slidePart, shape, rect) Then result.Add(rect)
        Next
        Return result
    End Function

    Private Shared Function PowerPointVisualRectsOverlap(a As AutoPilotPowerPointVisualRect, b As AutoPilotPowerPointVisualRect) As Boolean
        Return a.X < b.X + b.W AndAlso a.X + a.W > b.X AndAlso a.Y < b.Y + b.H AndAlso a.Y + a.H > b.Y
    End Function

    Private Shared Function GetOpenXmlPowerPointRichCanvas(slidePart As DocumentFormat.OpenXml.Packaging.SlidePart,
                                                            slideCx As Int64,
                                                            slideCy As Int64,
                                                            settings As Dictionary(Of String, String),
                                                            ByRef contentZones As List(Of AutoPilotPowerPointVisualRect)) As AutoPilotPowerPointVisualRect
        contentZones = New List(Of AutoPilotPowerPointVisualRect)()
        Dim layoutTree As DocumentFormat.OpenXml.Presentation.ShapeTree = slidePart?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree
        If layoutTree IsNot Nothing Then
            For Each shape As DocumentFormat.OpenXml.Presentation.Shape In layoutTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
                ' Only genuine placeholders define content geometry. Decorative/background layout shapes
                ' are never promoted to drawing zones, otherwise a large decoration can swallow the title area.
                If Not HasOpenXmlPowerPointPlaceholder(shape) Then Continue For
                Dim kind As String = GetOpenXmlPowerPointPlaceholderKind(shape)
                If IsOpenXmlPowerPointProtectedPlaceholder(kind) Then Continue For
                Dim rect As AutoPilotPowerPointVisualRect
                If Not TryResolveOpenXmlPowerPointPlaceholderRect(slidePart, shape, rect) Then Continue For
                If rect.W * rect.H < CLng(slideCx * slideCy * 0.05R) Then Continue For
                contentZones.Add(rect)
            Next
        End If

        Dim canvas As AutoPilotPowerPointVisualRect
        If contentZones.Count > 0 Then
            canvas.X = contentZones.Min(Function(r) r.X)
            canvas.Y = contentZones.Min(Function(r) r.Y)
            Dim maxX As Int64 = contentZones.Max(Function(r) r.X + r.W)
            Dim maxY As Int64 = contentZones.Max(Function(r) r.Y + r.H)
            canvas.W = maxX - canvas.X
            canvas.H = maxY - canvas.Y
        Else
            canvas.X = CLng(slideCx * 0.055R)
            canvas.Y = CLng(slideCy * 0.285R)
            canvas.W = slideCx - 2 * canvas.X
            canvas.H = slideCy - canvas.Y - CLng(slideCy * 0.10R)
        End If

        Dim insetPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.canvas_inset_pct", 2.0R)
        Dim insetX As Int64 = CLng(canvas.W * insetPct / 100.0R)
        Dim insetY As Int64 = CLng(canvas.H * insetPct / 100.0R)
        canvas.X += insetX
        canvas.Y += insetY
        canvas.W = Math.Max(1L, canvas.W - 2 * insetX)
        canvas.H = Math.Max(1L, canvas.H - 2 * insetY)
        Return canvas
    End Function

    Private Shared Function GetPowerPointGuidanceSettingString(settings As Dictionary(Of String, String), key As String, defaultValue As String) As String
        If settings Is Nothing OrElse Not settings.ContainsKey(key) Then Return defaultValue
        Dim value As String = If(settings(key), "").Trim()
        Return If(value = "", defaultValue, value)
    End Function

    Private Shared Function InsetPowerPointVisualRect(rect As AutoPilotPowerPointVisualRect, insetPct As Double) As AutoPilotPowerPointVisualRect
        Dim result As AutoPilotPowerPointVisualRect = rect
        Dim boundedPct As Double = Math.Max(0.0R, Math.Min(20.0R, insetPct))
        Dim insetX As Int64 = CLng(result.W * boundedPct / 100.0R)
        Dim insetY As Int64 = CLng(result.H * boundedPct / 100.0R)
        result.X += insetX
        result.Y += insetY
        result.W = Math.Max(1L, result.W - 2 * insetX)
        result.H = Math.Max(1L, result.H - 2 * insetY)
        Return result
    End Function

    Private Shared Function FitPowerPointVisualRectHeight(rect As AutoPilotPowerPointVisualRect, heightPct As Double) As AutoPilotPowerPointVisualRect
        Dim result As AutoPilotPowerPointVisualRect = rect
        Dim boundedPct As Double = Math.Max(35.0R, Math.Min(100.0R, heightPct))
        Dim targetH As Int64 = Math.Max(1L, CLng(result.H * boundedPct / 100.0R))
        result.Y += (result.H - targetH) \ 2
        result.H = targetH
        Return result
    End Function

    Private Shared Function FindOpenXmlPowerPointElementByLocalName(root As DocumentFormat.OpenXml.OpenXmlElement, localName As String) As DocumentFormat.OpenXml.OpenXmlElement
        If root Is Nothing OrElse String.IsNullOrWhiteSpace(localName) Then Return Nothing
        For Each child As DocumentFormat.OpenXml.OpenXmlElement In root.ChildElements
            If String.Equals(child.LocalName, localName, StringComparison.OrdinalIgnoreCase) Then Return child
            Dim nested As DocumentFormat.OpenXml.OpenXmlElement = FindOpenXmlPowerPointElementByLocalName(child, localName)
            If nested IsNot Nothing Then Return nested
        Next
        Return Nothing
    End Function

    Private Shared Function GetOpenXmlPowerPointThemePalette(slidePart As DocumentFormat.OpenXml.Packaging.SlidePart) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
        Try
            Dim themePart As DocumentFormat.OpenXml.Packaging.ThemePart = slidePart?.SlideLayoutPart?.SlideMasterPart?.ThemePart
            Dim themeRoot As DocumentFormat.OpenXml.OpenXmlElement = themePart?.Theme
            If themeRoot Is Nothing Then Return result
            Dim colorScheme As DocumentFormat.OpenXml.OpenXmlElement = FindOpenXmlPowerPointElementByLocalName(themeRoot, "clrScheme")
            If colorScheme Is Nothing Then Return result

            For Each schemeEntry As DocumentFormat.OpenXml.OpenXmlElement In colorScheme.ChildElements
                If schemeEntry Is Nothing OrElse schemeEntry.ChildElements.Count = 0 Then Continue For
                Dim colorNode As DocumentFormat.OpenXml.OpenXmlElement = schemeEntry.ChildElements(0)
                Dim raw As String = colorNode.GetAttribute("val", "").Value
                If String.IsNullOrWhiteSpace(raw) Then raw = colorNode.GetAttribute("lastClr", "").Value
                raw = If(raw, "").Trim().TrimStart("#"c)
                If System.Text.RegularExpressions.Regex.IsMatch(raw, "^[0-9A-Fa-f]{6}$") Then
                    result(schemeEntry.LocalName.ToLowerInvariant()) = raw.ToUpperInvariant()
                End If
            Next
        Catch ex As System.Exception
            Debug.WriteLine($"PowerPoint theme palette read failed: {ex.Message}")
        End Try
        Return result
    End Function

    Private Shared Function ResolveOpenXmlPowerPointPaletteColor(palette As Dictionary(Of String, String), slotName As String, fallbackHex As String) As String
        Dim key As String = If(slotName, "").Trim().ToLowerInvariant()
        If palette IsNot Nothing AndAlso key <> "" AndAlso palette.ContainsKey(key) Then
            Return NormalizeOpenXmlPowerPointHex(palette(key), fallbackHex)
        End If
        Return NormalizeOpenXmlPowerPointHex(fallbackHex, "17365D")
    End Function

    Private Shared Function MixOpenXmlPowerPointHex(colorHex As String, backgroundHex As String, backgroundPct As Double) As String
        Dim foreground As String = NormalizeOpenXmlPowerPointHex(colorHex, "17365D")
        Dim backgroundColor As String = NormalizeOpenXmlPowerPointHex(backgroundHex, "FFFFFF")
        Dim pct As Double = Math.Max(0.0R, Math.Min(100.0R, backgroundPct)) / 100.0R
        Dim inv As Double = 1.0R - pct
        Dim values As New List(Of Integer)()
        For offset As Integer = 0 To 4 Step 2
            Dim fg As Integer = System.Convert.ToInt32(foreground.Substring(offset, 2), 16)
            Dim bg As Integer = System.Convert.ToInt32(backgroundColor.Substring(offset, 2), 16)
            values.Add(CInt(Math.Round(fg * inv + bg * pct)))
        Next
        Return String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0:X2}{1:X2}{2:X2}", values(0), values(1), values(2))
    End Function

    Private Shared Function GetOpenXmlPowerPointPaletteSlots(settings As Dictionary(Of String, String)) As List(Of String)
        Dim raw As String = GetPowerPointGuidanceSettingString(settings, "rich.palette_slots", "accent1,accent3,accent2,accent5")
        Dim result As New List(Of String)()
        For Each item As String In raw.Split(","c)
            Dim value As String = item.Trim().ToLowerInvariant()
            If value <> "" Then result.Add(value)
        Next
        If result.Count = 0 Then result.AddRange(New String() {"accent1", "accent3", "accent2", "accent5"})
        Return result
    End Function

    Private Shared Function ResolveOpenXmlPowerPointToneSlot(settings As Dictionary(Of String, String), tone As String, fallbackSlot As String) As String
        Select Case If(tone, "").Trim().ToLowerInvariant()
            Case "positive", "success", "recommended"
                Return GetPowerPointGuidanceSettingString(settings, "rich.tone_positive_slot", fallbackSlot)
            Case "negative", "risk", "warning"
                Return GetPowerPointGuidanceSettingString(settings, "rich.tone_negative_slot", fallbackSlot)
            Case "neutral"
                Return GetPowerPointGuidanceSettingString(settings, "rich.tone_neutral_slot", fallbackSlot)
            Case Else
                Return fallbackSlot
        End Select
    End Function

    Private Shared Function EstimatePowerPointBodyFont(text As String,
                                                        boxW As Int64,
                                                        boxH As Int64,
                                                        slideCx As Int64,
                                                        slideCy As Int64,
                                                        preferredPt As Integer,
                                                        minimumPt As Integer) As Integer
        If String.IsNullOrWhiteSpace(text) Then Return preferredPt
        Dim widthRatio As Double = Math.Max(0.08R, boxW / CDbl(slideCx))
        Dim heightRatio As Double = Math.Max(0.08R, boxH / CDbl(slideCy))
        Dim capacity As Double = 1600.0R * widthRatio * heightRatio
        Dim density As Double = text.Length / Math.Max(1.0R, capacity)
        Dim result As Integer = preferredPt
        If density > 1.0R Then result -= CInt(Math.Ceiling((density - 1.0R) * 4.0R))
        Return Math.Max(minimumPt, Math.Min(preferredPt, result))
    End Function

    Private Shared Function PowerPointBodyTextFitsAtMinimum(text As String,
                                                               boxW As Int64,
                                                               boxH As Int64,
                                                               slideCx As Int64,
                                                               slideCy As Int64,
                                                               minimumPt As Integer) As Boolean
        If String.IsNullOrWhiteSpace(text) Then Return True
        Dim widthRatio As Double = Math.Max(0.08R, boxW / CDbl(slideCx))
        Dim heightRatio As Double = Math.Max(0.08R, boxH / CDbl(slideCy))
        Dim capacityAt15Pt As Double = 1600.0R * widthRatio * heightRatio
        Dim capacityAtMinimum As Double = capacityAt15Pt * (15.0R / Math.Max(1, minimumPt))
        Return text.Length <= capacityAtMinimum * 1.08R
    End Function

    Private Shared Sub EnsurePowerPointRichTextDensity(slideOrdinal As Integer,
                                                        semanticLayout As String,
                                                        itemOrdinal As Integer,
                                                        text As String,
                                                        boxW As Int64,
                                                        boxH As Int64,
                                                        slideCx As Int64,
                                                        slideCy As Int64,
                                                        minimumPt As Integer)
        If PowerPointBodyTextFitsAtMinimum(text, boxW, boxH, slideCx, slideCy, minimumPt) Then Exit Sub
        Throw New System.Exception($"PowerPoint rich-content density guard: slide {slideOrdinal} '{semanticLayout}' item {itemOrdinal} cannot fit at the configured minimum body font of {minimumPt} pt. Shorten the rich item, split the slide, or use a native text/data layout instead of shrinking or clipping text.")
    End Sub

    Private Shared Function CreateOpenXmlPowerPointConnector(shapeId As UInt32,
                                                              name As String,
                                                              x As Int64,
                                                              y As Int64,
                                                              w As Int64,
                                                              h As Int64,
                                                              colorHex As String) As DocumentFormat.OpenXml.Presentation.Shape
        Return CreateOpenXmlPowerPointTextShape(shapeId, name, x, y, Math.Max(w, 12000L), Math.Max(h, 12000L), colorHex, colorHex, colorHex, New List(Of Tuple(Of String, Integer, Boolean))())
    End Function

    Private Shared Function ApplyAutoPilotPowerPointRichContentOpenXml(outputPath As String,
                                                                        slidesArray As JArray,
                                                                        args As Dictionary(Of String, Object),
                                                                        design As AutoPilotDesignResolution,
                                                                        templatePath As String,
                                                                        context As ToolExecutionContext,
                                                                        ByRef errorText As String) As Boolean
        errorText = ""
        If String.IsNullOrWhiteSpace(outputPath) OrElse slidesArray Is Nothing Then Return True

        Try
            Dim guidance As String = ReadAutoPilotPowerPointDesignGuidance(design, templatePath, context)
            Dim settings As Dictionary(Of String, String) = ParsePowerPointGuidanceVisualSettings(guidance)
            Dim theme As JObject = GetPowerPointTheme(args)
            Dim argumentAccentHex As String = NormalizeOpenXmlPowerPointHex(theme.Value(Of String)("accent"), "17365D")
            Dim argumentSecondaryHex As String = NormalizeOpenXmlPowerPointHex(theme.Value(Of String)("secondary"), "2F75B5")
            Dim argumentTextHex As String = NormalizeOpenXmlPowerPointHex(theme.Value(Of String)("text"), "202124")
            Dim argumentLightHex As String = NormalizeOpenXmlPowerPointHex(theme.Value(Of String)("light"), "FFFFFF")
            Dim minBodyPt As Integer = GetPowerPointGuidanceSettingInteger(settings, "rich.min_body_font_pt", 15)
            Dim minTitlePt As Integer = GetPowerPointGuidanceSettingInteger(settings, "rich.min_title_font_pt", 18)
            Dim gapPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.gap_pct", 2.2R)
            Dim fillTintPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.fill_tint_pct", 90.0R)
            Dim altFillTintPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.alt_fill_tint_pct", 95.0R)
            Dim lineTintPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.line_tint_pct", 45.0R)
            Dim zoneInsetPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.zone_inset_pct", 1.5R)
            Dim paletteSlots As List(Of String) = GetOpenXmlPowerPointPaletteSlots(settings)

            Using document As DocumentFormat.OpenXml.Packaging.PresentationDocument =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, True)

                Dim presentationPart As DocumentFormat.OpenXml.Packaging.PresentationPart = document.PresentationPart
                Dim slideIds As List(Of DocumentFormat.OpenXml.Presentation.SlideId) =
                    presentationPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)().ToList()
                Dim firstGeneratedIndex As Integer = slideIds.Count - slidesArray.Count
                Dim slideCx As Int64 = CLng(presentationPart.Presentation.SlideSize.Cx.Value)
                Dim slideCy As Int64 = CLng(presentationPart.Presentation.SlideSize.Cy.Value)

                For ordinal As Integer = 1 To slidesArray.Count
                    Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
                    If slideObj Is Nothing Then Continue For
                    Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), ordinal, 0)
                    If semanticLayout <> "cards" AndAlso semanticLayout <> "comparison" AndAlso semanticLayout <> "structure" AndAlso semanticLayout <> "process" AndAlso semanticLayout <> "timeline" AndAlso semanticLayout <> "matrix" AndAlso semanticLayout <> "chart" AndAlso semanticLayout <> "kpi" Then Continue For

                    Dim slideId As DocumentFormat.OpenXml.Presentation.SlideId = slideIds(firstGeneratedIndex + ordinal - 1)
                    Dim slidePart As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presentationPart.GetPartById(slideId.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                    Dim shapeTree As DocumentFormat.OpenXml.Presentation.ShapeTree = slidePart?.Slide?.CommonSlideData?.ShapeTree
                    If shapeTree Is Nothing Then Throw New System.Exception($"Slide {ordinal} has no writable shape tree.")

                    Dim contentZones As List(Of AutoPilotPowerPointVisualRect) = Nothing
                    Dim canvas As AutoPilotPowerPointVisualRect = GetOpenXmlPowerPointRichCanvas(slidePart, slideCx, slideCy, settings, contentZones)
                    Dim palette As Dictionary(Of String, String) = GetOpenXmlPowerPointThemePalette(slidePart)
                    palette("accent1") = argumentAccentHex
                    palette("accent2") = argumentSecondaryHex
                    palette("dk1") = argumentTextHex
                    palette("lt1") = argumentLightHex
                    Dim textHex As String = ResolveOpenXmlPowerPointPaletteColor(palette, "dk1", argumentTextHex)
                    Dim lightHex As String = ResolveOpenXmlPowerPointPaletteColor(palette, "lt1", argumentLightHex)
                    Dim accentHex As String = ResolveOpenXmlPowerPointPaletteColor(palette, "accent1", argumentAccentHex)
                    Dim secondaryHex As String = ResolveOpenXmlPowerPointPaletteColor(palette, "accent2", argumentSecondaryHex)
                    Dim lineHex As String = MixOpenXmlPowerPointHex(accentHex, lightHex, lineTintPct)
                    Dim nextId As UInt32 = GetNextOpenXmlPowerPointShapeId(shapeTree)
                    Dim gap As Int64 = CLng(canvas.W * gapPct / 100.0R)
                    Dim added As Integer = 0

                    If semanticLayout = "cards" Then
                        Dim cards As JArray = TryCast(slideObj("cards"), JArray)
                        If cards Is Nothing OrElse cards.Count = 0 Then Continue For
                        Dim count As Integer = Math.Min(GetPowerPointGuidanceSettingInteger(settings, "rich.cards_max_items", 4), cards.Count)
                        Dim columns As Integer = If(count <= 2, count, 2)
                        Dim rows As Integer = CInt(Math.Ceiling(count / CDbl(columns)))
                        Dim cardsRect As AutoPilotPowerPointVisualRect = FitPowerPointVisualRectHeight(canvas, GetPowerPointGuidanceSettingDouble(settings, "rich.cards_height_pct", 86.0R))
                        Dim cellW As Int64 = (cardsRect.W - gap * (columns - 1)) \ columns
                        Dim cellH As Int64 = (cardsRect.H - gap * (rows - 1)) \ rows
                        For cardIndex As Integer = 0 To count - 1
                            Dim card As JObject = TryCast(cards(cardIndex), JObject)
                            If card Is Nothing Then Continue For
                            Dim col As Integer = cardIndex Mod columns
                            Dim row As Integer = cardIndex \ columns
                            Dim x As Int64 = cardsRect.X + col * (cellW + gap)
                            Dim y As Int64 = cardsRect.Y + row * (cellH + gap)
                            Dim bodyText As String = If(card.Value(Of String)("body"), "")
                            Dim cardDensityText As String = String.Join(vbLf, New String() {If(card.Value(Of String)("title"), ""), If(card.Value(Of String)("badge"), ""), bodyText})
                            EnsurePowerPointRichTextDensity(ordinal, semanticLayout, cardIndex + 1, cardDensityText, cellW, cellH, slideCx, slideCy, minBodyPt)
                            Dim bodyPt As Integer = EstimatePowerPointBodyFont(cardDensityText, cellW, cellH, slideCx, slideCy, 17, minBodyPt)
                            Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {
                                Tuple.Create(If(card.Value(Of String)("title"), ""), Math.Max(minTitlePt, bodyPt + 3), True),
                                Tuple.Create(If(card.Value(Of String)("badge"), ""), Math.Max(minBodyPt, bodyPt - 1), True),
                                Tuple.Create(bodyText, bodyPt, False)
                            }
                            Dim fallbackSlot As String = paletteSlots(cardIndex Mod paletteSlots.Count)
                            Dim toneSlot As String = ResolveOpenXmlPowerPointToneSlot(settings, card.Value(Of String)("tone"), fallbackSlot)
                            Dim cardAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, toneSlot, accentHex)
                            Dim cardFill As String = MixOpenXmlPowerPointHex(cardAccent, lightHex, If(cardIndex Mod 2 = 0, fillTintPct, altFillTintPct))
                            Dim cardLine As String = MixOpenXmlPowerPointHex(cardAccent, lightHex, lineTintPct)
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Card {cardIndex + 1}", x, y, cellW, cellH, cardFill, cardLine, textHex, paragraphs))
                            nextId += 1UI : added += 1
                            Dim accentStripH As Int64 = Math.Max(18000L, CLng(cellH * 0.018R))
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, $"Red Ink Card Accent {cardIndex + 1}", x, y, cellW, accentStripH, cardAccent))
                            nextId += 1UI : added += 1
                        Next

                    ElseIf semanticLayout = "comparison" Then
                        Dim comparison As JObject = TryCast(slideObj("comparison"), JObject)
                        Dim compareColumns As JArray = TryCast(comparison?("columns"), JArray)
                        If compareColumns Is Nothing OrElse compareColumns.Count = 0 Then Continue For
                        Dim count As Integer = Math.Min(4, compareColumns.Count)
                        Dim colRects As New List(Of AutoPilotPowerPointVisualRect)()
                        If contentZones IsNot Nothing AndAlso contentZones.Count >= count Then
                            For Each zone As AutoPilotPowerPointVisualRect In contentZones.OrderBy(Function(r) r.X).Take(count)
                                colRects.Add(InsetPowerPointVisualRect(zone, zoneInsetPct))
                            Next
                        Else
                            Dim colW As Int64 = (canvas.W - gap * (count - 1)) \ count
                            For i As Integer = 0 To count - 1
                                colRects.Add(New AutoPilotPowerPointVisualRect With {.X = canvas.X + i * (colW + gap), .Y = canvas.Y, .W = colW, .H = canvas.H})
                            Next
                        End If
                        For colIndex As Integer = 0 To count - 1
                            Dim column As JObject = TryCast(compareColumns(colIndex), JObject)
                            If column Is Nothing Then Continue For
                            Dim rect As AutoPilotPowerPointVisualRect = FitPowerPointVisualRectHeight(colRects(colIndex), GetPowerPointGuidanceSettingDouble(settings, "rich.comparison_height_pct", 90.0R))
                            Dim items As JArray = TryCast(column("items"), JArray)
                            Dim itemText As String = If(items Is Nothing, "", String.Join(vbLf, items.Select(Function(t) "• " & t.ToString())))
                            Dim comparisonDensityText As String = String.Join(vbLf, New String() {If(column.Value(Of String)("title"), ""), itemText, If(column.Value(Of String)("verdict"), "")})
                            EnsurePowerPointRichTextDensity(ordinal, semanticLayout, colIndex + 1, comparisonDensityText, rect.W, rect.H, slideCx, slideCy, minBodyPt)
                            Dim bodyPt As Integer = EstimatePowerPointBodyFont(comparisonDensityText, rect.W, rect.H, slideCx, slideCy, 16, minBodyPt)
                            Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(If(column.Value(Of String)("title"), ""), Math.Max(minTitlePt, bodyPt + 3), True)}
                            If items IsNot Nothing Then
                                For Each item As JToken In items
                                    paragraphs.Add(Tuple.Create("• " & item.ToString(), bodyPt, False))
                                Next
                            End If
                            paragraphs.Add(Tuple.Create(If(column.Value(Of String)("verdict"), ""), Math.Max(minBodyPt, bodyPt), True))
                            Dim fallbackSlot As String = paletteSlots(colIndex Mod paletteSlots.Count)
                            Dim toneSlot As String = ResolveOpenXmlPowerPointToneSlot(settings, column.Value(Of String)("tone"), fallbackSlot)
                            Dim columnAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, toneSlot, accentHex)
                            Dim fill As String = MixOpenXmlPowerPointHex(columnAccent, lightHex, If(colIndex Mod 2 = 0, fillTintPct, altFillTintPct))
                            Dim columnLine As String = MixOpenXmlPowerPointHex(columnAccent, lightHex, lineTintPct)
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Comparison {colIndex + 1}", rect.X, rect.Y, rect.W, rect.H, fill, columnLine, textHex, paragraphs))
                            nextId += 1UI : added += 1
                            Dim comparisonAccentH As Int64 = Math.Max(18000L, CLng(rect.H * 0.015R))
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, $"Red Ink Comparison Accent {colIndex + 1}", rect.X, rect.Y, rect.W, comparisonAccentH, columnAccent))
                            nextId += 1UI : added += 1
                        Next

                    ElseIf semanticLayout = "structure" Then
                        Dim structureObj As JObject = TryCast(slideObj("structure"), JObject)
                        If structureObj Is Nothing Then Continue For
                        Dim topNode As JObject = TryCast(structureObj("top"), JObject)
                        Dim children As JArray = TryCast(structureObj("children"), JArray)
                        Dim parentWidthPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.structure_parent_width_pct", 44.0R)
                        Dim topW As Int64 = CLng(canvas.W * parentWidthPct / 100.0R)
                        Dim topHeightPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.structure_parent_height_pct", 30.0R)
                        Dim topH As Int64 = CLng(canvas.H * Math.Max(20.0R, Math.Min(40.0R, topHeightPct)) / 100.0R)
                        Dim topX As Int64 = canvas.X + (canvas.W - topW) \ 2
                        Dim topY As Int64 = canvas.Y + CLng(canvas.H * 0.03R)
                        If children IsNot Nothing AndAlso children.Count > 0 Then
                            Dim childCount As Integer = Math.Min(4, children.Count)
                            Dim childTopPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.structure_child_top_pct", 60.0R)
                            Dim childHeightPct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.structure_child_height_pct", 34.0R)
                            Dim childY As Int64 = canvas.Y + CLng(canvas.H * Math.Max(48.0R, Math.Min(72.0R, childTopPct)) / 100.0R)
                            Dim childH As Int64 = Math.Min(canvas.Y + canvas.H - childY, CLng(canvas.H * Math.Max(22.0R, Math.Min(42.0R, childHeightPct)) / 100.0R))
                            Dim maxSinglePct As Double = GetPowerPointGuidanceSettingDouble(settings, "rich.structure_single_child_width_pct", 48.0R)
                            Dim childW As Int64 = If(childCount = 1, CLng(canvas.W * maxSinglePct / 100.0R), (canvas.W - gap * (childCount - 1)) \ childCount)
                            Dim totalChildrenW As Int64 = childW * childCount + gap * (childCount - 1)
                            Dim startX As Int64 = canvas.X + (canvas.W - totalChildrenW) \ 2
                            Dim parentCenterX As Int64 = topX + topW \ 2
                            Dim midY As Int64 = topY + topH + (childY - (topY + topH)) \ 2
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, "Red Ink Structure Stem", parentCenterX - 6000, topY + topH, 12000, midY - (topY + topH), lineHex))
                            nextId += 1UI : added += 1
                            If childCount > 1 Then
                                Dim firstCenter As Int64 = startX + childW \ 2
                                Dim lastCenter As Int64 = startX + (childCount - 1) * (childW + gap) + childW \ 2
                                shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, "Red Ink Structure Rail", firstCenter, midY - 6000, lastCenter - firstCenter, 12000, lineHex))
                                nextId += 1UI : added += 1
                            End If
                            For childIndex As Integer = 0 To childCount - 1
                                Dim child As JObject = TryCast(children(childIndex), JObject)
                                If child Is Nothing Then Continue For
                                Dim childX As Int64 = startX + childIndex * (childW + gap)
                                Dim childCenter As Int64 = childX + childW \ 2
                                shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, $"Red Ink Structure Branch {childIndex + 1}", childCenter - 6000, midY, 12000, childY - midY, lineHex))
                                nextId += 1UI : added += 1
                                Dim bodyText As String = If(child.Value(Of String)("body"), "")
                                Dim childDensityText As String = String.Join(vbLf, New String() {If(child.Value(Of String)("title"), ""), bodyText})
                                EnsurePowerPointRichTextDensity(ordinal, semanticLayout, childIndex + 1, childDensityText, childW, childH, slideCx, slideCy, minBodyPt)
                                Dim bodyPt As Integer = EstimatePowerPointBodyFont(childDensityText, childW, childH, slideCx, slideCy, 16, minBodyPt)
                                Dim childParagraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(If(child.Value(Of String)("title"), ""), Math.Max(minTitlePt, bodyPt + 3), True), Tuple.Create(bodyText, bodyPt, False)}
                                Dim childAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots((childIndex + 1) Mod paletteSlots.Count), secondaryHex)
                                Dim childFill As String = MixOpenXmlPowerPointHex(childAccent, lightHex, altFillTintPct)
                                Dim childLine As String = MixOpenXmlPowerPointHex(childAccent, lightHex, lineTintPct)
                                shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Structure Child {childIndex + 1}", childX, childY, childW, childH, childFill, childLine, textHex, childParagraphs))
                                nextId += 1UI : added += 1
                            Next
                        End If
                        If topNode IsNot Nothing Then
                            Dim topBody As String = If(topNode.Value(Of String)("body"), "")
                            Dim topDensityText As String = String.Join(vbLf, New String() {If(topNode.Value(Of String)("title"), ""), topBody})
                            EnsurePowerPointRichTextDensity(ordinal, semanticLayout, 0, topDensityText, topW, topH, slideCx, slideCy, minBodyPt)
                            Dim topBodyPt As Integer = EstimatePowerPointBodyFont(topDensityText, topW, topH, slideCx, slideCy, 16, minBodyPt)
                            Dim topParagraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(If(topNode.Value(Of String)("title"), ""), Math.Max(minTitlePt, topBodyPt + 3), True), Tuple.Create(topBody, topBodyPt, False)}
                            Dim topFill As String = MixOpenXmlPowerPointHex(accentHex, lightHex, fillTintPct)
                            Dim topLine As String = MixOpenXmlPowerPointHex(accentHex, lightHex, lineTintPct)
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, "Red Ink Structure Top", topX, topY, topW, topH, topFill, topLine, textHex, topParagraphs))
                            nextId += 1UI : added += 1
                        End If

                    ElseIf semanticLayout = "process" OrElse semanticLayout = "timeline" Then
                        Dim steps As JArray = If(semanticLayout = "process", TryCast(slideObj("steps"), JArray), TryCast(slideObj("events"), JArray))
                        If steps Is Nothing AndAlso semanticLayout = "timeline" Then steps = TryCast(slideObj("timeline"), JArray)
                        If steps Is Nothing AndAlso semanticLayout = "timeline" Then
                            Dim timelineObject As JObject = TryCast(slideObj("timeline"), JObject)
                            If timelineObject IsNot Nothing Then steps = TryCast(timelineObject("events"), JArray)
                        End If
                        If steps Is Nothing OrElse steps.Count = 0 Then Continue For

                        Dim count As Integer = Math.Min(GetPowerPointGuidanceSettingInteger(settings, "rich.process_max_items", 4), steps.Count)
                        Dim processRect As AutoPilotPowerPointVisualRect = FitPowerPointVisualRectHeight(canvas, GetPowerPointGuidanceSettingDouble(settings, "rich.process_height_pct", 80.0R))
                        Dim segmentW As Int64 = processRect.W \ count
                        Dim lineY As Int64 = processRect.Y + CLng(processRect.H * 0.24R)
                        Dim firstCenterX As Int64 = processRect.X + segmentW \ 2
                        Dim lastCenterX As Int64 = processRect.X + (count - 1) * segmentW + segmentW \ 2
                        Dim timelineLineH As Int64 = Math.Max(16000L, CLng(processRect.H * 0.008R))
                        shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, $"Red Ink {semanticLayout} Line", firstCenterX, lineY - timelineLineH \ 2, Math.Max(12000L, lastCenterX - firstCenterX), timelineLineH, lineHex))
                        nextId += 1UI : added += 1

                        Dim markerSize As Int64 = Math.Max(90000L, Math.Min(150000L, CLng(processRect.H * 0.07R)))
                        Dim cardTop As Int64 = lineY + CLng(processRect.H * 0.12R)
                        Dim cardH As Int64 = Math.Max(1L, processRect.Y + processRect.H - cardTop)
                        Dim cardW As Int64 = Math.Max(1L, segmentW - gap)

                        For i As Integer = 0 To count - 1
                            Dim item As JObject = TryCast(steps(i), JObject)
                            If item Is Nothing Then Continue For
                            Dim centerX As Int64 = processRect.X + i * segmentW + segmentW \ 2
                            Dim x As Int64 = processRect.X + i * segmentW + gap \ 2
                            Dim label As String = If(item.Value(Of String)("label"), If(item.Value(Of String)("date"), If(item.Value(Of String)("phase"), (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture))))
                            Dim bodyText As String = If(item.Value(Of String)("body"), "")
                            Dim titleText As String = If(item.Value(Of String)("title"), "")
                            Dim processDensityText As String = String.Join(vbLf, New String() {label, titleText, bodyText})
                            EnsurePowerPointRichTextDensity(ordinal, semanticLayout, i + 1, processDensityText, cardW, cardH, slideCx, slideCy, minBodyPt)
                            Dim bodyPt As Integer = EstimatePowerPointBodyFont(processDensityText, cardW, cardH, slideCx, slideCy, 16, minBodyPt)
                            Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {
                                Tuple.Create(label, Math.Max(minBodyPt, bodyPt), True),
                                Tuple.Create(titleText, Math.Max(minTitlePt, bodyPt + 3), True),
                                Tuple.Create(bodyText, bodyPt, False)
                            }
                            Dim stepAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots(i Mod paletteSlots.Count), secondaryHex)
                            Dim stepFill As String = MixOpenXmlPowerPointHex(stepAccent, lightHex, If(i Mod 2 = 0, fillTintPct, altFillTintPct))
                            Dim stepLine As String = MixOpenXmlPowerPointHex(stepAccent, lightHex, lineTintPct)
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, $"Red Ink {semanticLayout} Marker {i + 1}", centerX - markerSize \ 2, lineY - markerSize \ 2, markerSize, markerSize, stepAccent))
                            nextId += 1UI : added += 1
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink {semanticLayout} {i + 1}", x, cardTop, cardW, cardH, stepFill, stepLine, textHex, paragraphs))
                            nextId += 1UI : added += 1
                        Next

                    ElseIf semanticLayout = "kpi" Then
                        Dim kpis As JArray = TryCast(slideObj("kpis"), JArray)
                        If kpis Is Nothing OrElse kpis.Count = 0 Then Continue For
                        Dim count As Integer = Math.Min(4, kpis.Count)
                        Dim colW As Int64 = (canvas.W - gap * (count - 1)) \ count
                        Dim boxH As Int64 = CLng(canvas.H * 0.55R)
                        Dim y As Int64 = canvas.Y + (canvas.H - boxH) \ 2
                        For i As Integer = 0 To count - 1
                            Dim kpi As JObject = TryCast(kpis(i), JObject)
                            If kpi Is Nothing Then Continue For
                            Dim x As Int64 = canvas.X + i * (colW + gap)
                            Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(If(kpi.Value(Of String)("value"), ""), 26, True), Tuple.Create(If(kpi.Value(Of String)("label"), ""), minTitlePt, True), Tuple.Create(If(kpi.Value(Of String)("detail"), ""), minBodyPt, False)}
                            Dim kpiAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots(i Mod paletteSlots.Count), accentHex)
                            Dim kpiFill As String = MixOpenXmlPowerPointHex(kpiAccent, lightHex, fillTintPct)
                            Dim kpiLine As String = MixOpenXmlPowerPointHex(kpiAccent, lightHex, lineTintPct)
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink KPI {i + 1}", x, y, colW, boxH, kpiFill, kpiLine, textHex, paragraphs))
                            nextId += 1UI : added += 1
                        Next

                    ElseIf semanticLayout = "matrix" Then
                        Dim matrix As JObject = TryCast(slideObj("matrix"), JObject)
                        If matrix Is Nothing Then Continue For
                        Dim items As JArray = TryCast(matrix("items"), JArray)
                        If items Is Nothing Then items = TryCast(matrix("points"), JArray)
                        Dim quadrants As JArray = TryCast(matrix("quadrants"), JArray)
                        If items Is Nothing AndAlso quadrants IsNot Nothing Then
                            Dim quadrantW As Int64 = (canvas.W - gap) \ 2
                            Dim quadrantH As Int64 = (canvas.H - gap) \ 2
                            For quadrantIndex As Integer = 0 To Math.Min(4, quadrants.Count) - 1
                                Dim quadrant As JObject = TryCast(quadrants(quadrantIndex), JObject)
                                If quadrant Is Nothing Then Continue For
                                Dim quadrantCol As Integer = quadrantIndex Mod 2
                                Dim quadrantRow As Integer = quadrantIndex \ 2
                                Dim quadrantX As Int64 = canvas.X + quadrantCol * (quadrantW + gap)
                                Dim quadrantY As Int64 = canvas.Y + quadrantRow * (quadrantH + gap)
                                Dim quadrantParagraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {
                                    Tuple.Create(If(quadrant.Value(Of String)("title"), ""), minTitlePt, True),
                                    Tuple.Create(If(quadrant.Value(Of String)("body"), ""), minBodyPt, False)
                                }
                                Dim quadrantAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots(quadrantIndex Mod paletteSlots.Count), accentHex)
                                Dim quadrantFill As String = MixOpenXmlPowerPointHex(quadrantAccent, lightHex, If(quadrantIndex Mod 2 = 0, fillTintPct, altFillTintPct))
                                Dim quadrantLine As String = MixOpenXmlPowerPointHex(quadrantAccent, lightHex, lineTintPct)
                                shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Matrix Quadrant {quadrantIndex + 1}", quadrantX, quadrantY, quadrantW, quadrantH, quadrantFill, quadrantLine, textHex, quadrantParagraphs))
                                nextId += 1UI : added += 1
                            Next
                        Else
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, "Red Ink Matrix X", canvas.X, canvas.Y + canvas.H - 24000, canvas.W, 18000, lineHex)) : nextId += 1UI : added += 1
                            shapeTree.Append(CreateOpenXmlPowerPointConnector(nextId, "Red Ink Matrix Y", canvas.X, canvas.Y, 18000, canvas.H, lineHex)) : nextId += 1UI : added += 1
                        End If
                        If items IsNot Nothing Then
                            Dim pointW As Int64 = CLng(canvas.W * 0.22R), pointH As Int64 = CLng(canvas.H * 0.20R)
                            For i As Integer = 0 To Math.Min(5, items.Count) - 1
                                Dim point As JObject = TryCast(items(i), JObject)
                                If point Is Nothing Then Continue For
                                Dim px As Double = Math.Max(0.0R, Math.Min(1.0R, If(point.Value(Of Double?)("x").HasValue, point.Value(Of Double?)("x").Value, (i + 1) / CDbl(items.Count + 1))))
                                Dim py As Double = Math.Max(0.0R, Math.Min(1.0R, If(point.Value(Of Double?)("y").HasValue, point.Value(Of Double?)("y").Value, 0.5R)))
                                Dim x As Int64 = canvas.X + CLng(px * Math.Max(1L, canvas.W - pointW)), y As Int64 = canvas.Y + CLng((1.0R - py) * Math.Max(1L, canvas.H - pointH))
                                Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(If(point.Value(Of String)("title"), If(point.Value(Of String)("label"), "")), minTitlePt, True), Tuple.Create(If(point.Value(Of String)("body"), ""), minBodyPt, False)}
                                Dim pointAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots(i Mod paletteSlots.Count), accentHex)
                                Dim pointFill As String = MixOpenXmlPowerPointHex(pointAccent, lightHex, fillTintPct)
                                Dim pointLine As String = MixOpenXmlPowerPointHex(pointAccent, lightHex, lineTintPct)
                                shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Matrix Point {i + 1}", x, y, pointW, pointH, pointFill, pointLine, textHex, paragraphs))
                                nextId += 1UI : added += 1
                            Next
                        End If

                    ElseIf semanticLayout = "chart" Then
                        Dim chart As JObject = TryCast(slideObj("chart"), JObject)
                        Dim values As JArray = TryCast(chart?("values"), JArray)
                        Dim labels As JArray = TryCast(chart?("labels"), JArray)
                        If labels Is Nothing Then labels = TryCast(chart?("categories"), JArray)
                        If values Is Nothing Then values = TryCast(chart?("data"), JArray)
                        If values Is Nothing Then
                            Dim series As JArray = TryCast(chart?("series"), JArray)
                            If series IsNot Nothing AndAlso series.Count > 0 Then
                                Dim firstSeries As JObject = TryCast(series(0), JObject)
                                If firstSeries IsNot Nothing Then values = TryCast(firstSeries("values"), JArray)
                            End If
                        End If
                        If values Is Nothing OrElse values.Count = 0 Then Continue For
                        Dim numeric As New List(Of Double)()
                        For Each token As JToken In values
                            Dim numericValue As Double
                            If Double.TryParse(token.ToString(), numericValue) Then numeric.Add(numericValue)
                        Next
                        If numeric.Count = 0 Then Continue For
                        Dim maxValue As Double = Math.Max(1.0R, numeric.Max())
                        Dim count As Integer = numeric.Count
                        Dim barGap As Int64 = CLng(canvas.W * 0.025R)
                        Dim barW As Int64 = (canvas.W - barGap * (count - 1)) \ count
                        For i As Integer = 0 To count - 1
                            Dim barH As Int64 = CLng(canvas.H * 0.72R * numeric(i) / maxValue)
                            Dim x As Int64 = canvas.X + i * (barW + barGap), y As Int64 = canvas.Y + canvas.H - barH
                            Dim label As String = If(labels IsNot Nothing AndAlso i < labels.Count, labels(i).ToString(), "")
                            Dim paragraphs As New List(Of Tuple(Of String, Integer, Boolean)) From {Tuple.Create(numeric(i).ToString("0.##"), minTitlePt, True), Tuple.Create(label, minBodyPt, False)}
                            Dim barAccent As String = ResolveOpenXmlPowerPointPaletteColor(palette, paletteSlots(i Mod paletteSlots.Count), If(i Mod 2 = 0, accentHex, secondaryHex))
                            shapeTree.Append(CreateOpenXmlPowerPointTextShape(nextId, $"Red Ink Chart Bar {i + 1}", x, y, barW, barH, barAccent, barAccent, lightHex, paragraphs))
                            nextId += 1UI : added += 1
                        Next
                    End If

                    If added = 0 Then Throw New System.Exception($"Slide {ordinal} required rich OpenXML content but no shapes were produced.")
                    slidePart.Slide.Save()
                    If context IsNot Nothing Then context.Log($"PowerPoint layout-aware rich OpenXML compositor applied: slide={ordinal}; semantic='{semanticLayout}'; shapes={added}; canvas={canvas.W}x{canvas.H}; zones={If(contentZones Is Nothing, 0, contentZones.Count)}.", "diag")
                Next
            End Using
        Catch ex As System.Exception
            errorText = $"PowerPoint rich-content OpenXML rendering failed: {ex.Message}"
            Return False
        End Try
        Return True
    End Function

    Private Shared Function NormalizeAutoPilotPowerPointSemanticLayout(layoutValue As String,
                                                                        slideIndex As Integer,
                                                                        existingSlideCount As Integer) As String
        Dim layout As String = If(layoutValue, "").Trim().ToLowerInvariant()
        If layout = "" Then
            If existingSlideCount = 0 AndAlso slideIndex = 1 Then
                layout = "title"
            Else
                layout = "bullets"
            End If
        End If

        Select Case layout
            Case "kpis" : Return "kpi"
            Case "org", "organization" : Return "structure"
            Case Else : Return layout
        End Select
    End Function

    Private Shared Function NormalizePowerPointLayoutKey(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim sb As New System.Text.StringBuilder()
        For Each ch As Char In value.Trim().ToLowerInvariant()
            If Char.IsLetterOrDigit(ch) Then sb.Append(ch)
        Next
        Return sb.ToString()
    End Function

    Private Shared Sub ResolveConfiguredPowerPointLayoutTarget(design As AutoPilotDesignResolution,
                                                               semanticLayout As String,
                                                               ByRef masterName As String,
                                                               ByRef layoutName As String)
        masterName = ""
        layoutName = ""
        If design Is Nothing OrElse design.ApplicationConfig Is Nothing Then Return

        Dim map As JObject = TryCast(design.ApplicationConfig("layout_map"), JObject)
        If map Is Nothing Then Return

        Dim target As JToken = map(semanticLayout)
        If target Is Nothing OrElse target.Type = JTokenType.Null Then Return

        If target.Type = JTokenType.String Then
            layoutName = target.ToString().Trim()
            Return
        End If

        Dim targetObject As JObject = TryCast(target, JObject)
        If targetObject Is Nothing Then Return
        masterName = If(targetObject.Value(Of String)("master"), "").Trim()
        layoutName = If(targetObject.Value(Of String)("layout"), "").Trim()
    End Sub

    Private Shared Function IsSyntheticPowerPointMasterName(value As String) As Boolean
        If String.IsNullOrWhiteSpace(value) Then Return True
        Dim trimmed As String = value.Trim()
        If System.Text.RegularExpressions.Regex.IsMatch(trimmed, "^Master\s+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then Return True
        Return trimmed.StartsWith("/ppt/slideMasters/", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function GetOpenXmlPowerPointMasterName(masterPart As DocumentFormat.OpenXml.Packaging.SlideMasterPart,
                                                            masterIndex As Integer) As String
        Dim value As String = ""
        Try
            value = If(masterPart?.SlideMaster?.CommonSlideData?.Name?.Value, "").Trim()
        Catch ex As System.Exception
        End Try

        ' PowerPoint often leaves p:cSld/@name empty on a slide master. In that
        ' case use the actual theme name carried by the master rather than assuming
        ' that a numeric master position has semantic meaning.
        If value = "" Then
            Try
                value = If(masterPart?.ThemePart?.Theme?.Name?.Value, "").Trim()
            Catch ex As System.Exception
            End Try
        End If

        ' Keep the relationship-backed part URI as the final stable metadata
        ' fallback. It is useful for diagnostics/LLM disambiguation, but is treated
        ' as an unknown COM master name during live layout binding.
        If value = "" Then
            Try
                value = If(masterPart?.Uri?.ToString(), "").Trim()
            Catch ex As System.Exception
            End Try
        End If
        If value = "" Then value = $"Master {masterIndex}"
        Return value
    End Function

    Private Shared Function GetOpenXmlPowerPointLayoutName(layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
                                                            layoutIndex As Integer) As String
        Dim value As String = ""
        Try
            value = If(layoutPart?.SlideLayout?.CommonSlideData?.Name?.Value, "").Trim()
        Catch ex As System.Exception
        End Try
        If value = "" Then value = $"Layout {layoutIndex}"
        Return value
    End Function

    Private Shared Function GetPowerPointComPlaceholderType(openXmlType As String) As Integer
        Dim key As String = If(openXmlType, "").Trim().ToLowerInvariant()
        Select Case key
            Case "", "implicit", "body" : Return 2
            Case "title" : Return 1
            Case "centertitle", "centeredtitle" : Return 3
            Case "subtitle" : Return 4
            Case "verticaltitle" : Return 5
            Case "verticalbody" : Return 6
            Case "object" : Return 7
            Case "chart" : Return 8
            Case "bitmap", "media" : Return 9
            Case "orgchart" : Return 11
            Case "table" : Return 12
            Case "slidenumber" : Return 13
            Case "header" : Return 14
            Case "footer" : Return 15
            Case "dateandtime" : Return 16
            Case "verticalobject" : Return 17
            Case "picture" : Return 18
            Case Else : Return 0
        End Select
    End Function

    Private Shared Function IsPowerPointTextPlaceholderType(openXmlType As String) As Boolean
        Dim comType As Integer = GetPowerPointComPlaceholderType(openXmlType)
        Return comType = 1 OrElse
               comType = 2 OrElse
               comType = 3 OrElse
               comType = 4 OrElse
               comType = 5 OrElse
               comType = 6 OrElse
               comType = 7 OrElse
               comType = 17
    End Function

    Private Shared Function BuildPowerPointPlaceholderDetail(shapeName As String,
                                                              placeholder As DocumentFormat.OpenXml.Presentation.PlaceholderShape,
                                                              transform As DocumentFormat.OpenXml.Drawing.Transform2D,
                                                              occurrence As Integer) As JObject
        Dim detail As New JObject()
        Dim typeName As String = If(placeholder?.Type?.Value.ToString(), "Body")
        If String.IsNullOrWhiteSpace(typeName) Then typeName = "Body"
        Dim comType As Integer = GetPowerPointComPlaceholderType(typeName)

        detail("name") = If(shapeName, "")
        detail("type") = typeName
        detail("occurrence") = occurrence
        detail("placeholder_key") = typeName.Trim().ToLowerInvariant() & ":" & occurrence.ToString(Globalization.CultureInfo.InvariantCulture)
        detail("com_type") = comType
        detail("text_capable") = IsPowerPointTextPlaceholderType(typeName)
        detail("index") = If(placeholder?.Index IsNot Nothing, CLng(placeholder.Index.Value), 0L)

        If transform IsNot Nothing Then
            If transform.Offset IsNot Nothing Then
                If transform.Offset.X IsNot Nothing Then detail("x") = CLng(transform.Offset.X.Value)
                If transform.Offset.Y IsNot Nothing Then detail("y") = CLng(transform.Offset.Y.Value)
            End If
            If transform.Extents IsNot Nothing Then
                If transform.Extents.Cx IsNot Nothing Then detail("width") = CLng(transform.Extents.Cx.Value)
                If transform.Extents.Cy IsNot Nothing Then detail("height") = CLng(transform.Extents.Cy.Value)
            End If
        End If
        Return detail
    End Function

    Private Shared Function InspectAutoPilotPowerPointTemplateLayouts(templatePath As String) As List(Of AutoPilotPowerPointLayoutCandidate)
        Dim result As New List(Of AutoPilotPowerPointLayoutCandidate)()
        If String.IsNullOrWhiteSpace(templatePath) OrElse Not System.IO.File.Exists(templatePath) Then Return result

        Try
            Using document As DocumentFormat.OpenXml.Packaging.PresentationDocument =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(templatePath, False)

                Dim presentationPart As DocumentFormat.OpenXml.Packaging.PresentationPart = document.PresentationPart
                If presentationPart Is Nothing Then Return result
                Dim templateSlideWidth As Long = 0L
                Dim templateSlideHeight As Long = 0L
                Try
                    If presentationPart.Presentation?.SlideSize IsNot Nothing Then
                        templateSlideWidth = CLng(presentationPart.Presentation.SlideSize.Cx.Value)
                        templateSlideHeight = CLng(presentationPart.Presentation.SlideSize.Cy.Value)
                    End If
                Catch ex As System.Exception
                End Try

                Dim byLayoutUri As New Dictionary(Of String, AutoPilotPowerPointLayoutCandidate)(StringComparer.OrdinalIgnoreCase)
                Dim masterIndex As Integer = 0
                Dim masterIdList As DocumentFormat.OpenXml.Presentation.SlideMasterIdList = presentationPart.Presentation?.SlideMasterIdList
                If masterIdList Is Nothing Then Return result

                For Each masterId As DocumentFormat.OpenXml.Presentation.SlideMasterId In masterIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideMasterId)()
                    If masterId.RelationshipId Is Nothing Then Continue For
                    Dim masterPart As DocumentFormat.OpenXml.Packaging.SlideMasterPart =
                        TryCast(presentationPart.GetPartById(masterId.RelationshipId), DocumentFormat.OpenXml.Packaging.SlideMasterPart)
                    If masterPart Is Nothing Then Continue For

                    masterIndex += 1
                    Dim masterName As String = GetOpenXmlPowerPointMasterName(masterPart, masterIndex)
                    Dim layoutIndex As Integer = 0
                    Dim layoutIdList As DocumentFormat.OpenXml.Presentation.SlideLayoutIdList = masterPart.SlideMaster?.SlideLayoutIdList
                    If layoutIdList Is Nothing Then Continue For

                    For Each layoutId As DocumentFormat.OpenXml.Presentation.SlideLayoutId In layoutIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideLayoutId)()
                        If layoutId.RelationshipId Is Nothing Then Continue For
                        Dim layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart =
                            TryCast(masterPart.GetPartById(layoutId.RelationshipId), DocumentFormat.OpenXml.Packaging.SlideLayoutPart)
                        If layoutPart Is Nothing Then Continue For

                        layoutIndex += 1
                        Dim candidate As New AutoPilotPowerPointLayoutCandidate() With {
                            .DesignName = masterName,
                            .LayoutName = GetOpenXmlPowerPointLayoutName(layoutPart, layoutIndex),
                            .SelectionReason = "template metadata",
                            .MasterOrdinal = masterIndex,
                            .LayoutOrdinal = layoutIndex,
                            .SlideWidth = templateSlideWidth,
                            .SlideHeight = templateSlideHeight
                        }

                        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree
                        If tree IsNot Nothing Then
                            Dim placeholderOccurrences As New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)

                            For Each shape As DocumentFormat.OpenXml.Presentation.Shape In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
                                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                If ph Is Nothing Then Continue For
                                Dim shapeName As String = If(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, "")
                                Dim typeName As String = If(ph.Type?.Value.ToString(), "Body")
                                If String.IsNullOrWhiteSpace(typeName) Then typeName = "Body"
                                Dim occurrence As Integer = If(placeholderOccurrences.ContainsKey(typeName), placeholderOccurrences(typeName), 0)
                                placeholderOccurrences(typeName) = occurrence + 1
                                Dim detail As JObject = BuildPowerPointPlaceholderDetail(shapeName, ph, shape.ShapeProperties?.Transform2D, occurrence)
                                Dim sampleText As String = String.Join(" ", shape.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().Select(Function(t) t.Text)).Trim()
                                If sampleText <> "" Then detail("sample_text") = sampleText
                                candidate.PlaceholderDetails.Add(detail)
                            Next

                            For Each picture As DocumentFormat.OpenXml.Presentation.Picture In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Picture)()
                                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                If ph Is Nothing Then Continue For
                                Dim shapeName As String = If(picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value, "")
                                Dim typeName As String = If(ph.Type?.Value.ToString(), "Picture")
                                If String.IsNullOrWhiteSpace(typeName) Then typeName = "Picture"
                                Dim occurrence As Integer = If(placeholderOccurrences.ContainsKey(typeName), placeholderOccurrences(typeName), 0)
                                placeholderOccurrences(typeName) = occurrence + 1
                                candidate.PlaceholderDetails.Add(BuildPowerPointPlaceholderDetail(shapeName, ph, picture.ShapeProperties?.Transform2D, occurrence))
                            Next
                        End If

                        result.Add(candidate)
                        If layoutPart.Uri IsNot Nothing Then byLayoutUri(layoutPart.Uri.ToString()) = candidate
                    Next
                Next

                Dim slideIdList As DocumentFormat.OpenXml.Presentation.SlideIdList = presentationPart.Presentation?.SlideIdList
                If slideIdList IsNot Nothing Then
                    For Each slideId As DocumentFormat.OpenXml.Presentation.SlideId In slideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                        If slideId.RelationshipId Is Nothing Then Continue For
                        Try
                            Dim slidePart As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presentationPart.GetPartById(slideId.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                            Dim layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = slidePart?.SlideLayoutPart
                            If layoutPart?.Uri Is Nothing Then Continue For
                            Dim key As String = layoutPart.Uri.ToString()
                            If Not byLayoutUri.ContainsKey(key) Then Continue For
                            Dim candidate As AutoPilotPowerPointLayoutCandidate = byLayoutUri(key)
                            If candidate.SampleSlides.Count >= 3 Then Continue For
                            Dim text As String = String.Join(" ", slidePart.Slide.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().Select(Function(t) t.Text)).Trim()
                            If text.Length > 700 Then text = text.Substring(0, 700) & "…"
                            If text <> "" Then candidate.SampleSlides.Add(text)
                        Catch ex As System.Exception
                        End Try
                    Next
                End If
            End Using
        Catch ex As System.Exception
        End Try

        Return result
    End Function

    Private Shared Function FindPowerPointCatalogLayout(catalog As IEnumerable(Of AutoPilotPowerPointLayoutCandidate),
                                                        masterName As String,
                                                        layoutName As String) As AutoPilotPowerPointLayoutCandidate
        If catalog Is Nothing OrElse String.IsNullOrWhiteSpace(layoutName) Then Return Nothing
        Dim layoutKey As String = NormalizePowerPointLayoutKey(layoutName)
        Dim matches As List(Of AutoPilotPowerPointLayoutCandidate) = catalog.
            Where(Function(c) c IsNot Nothing AndAlso NormalizePowerPointLayoutKey(c.LayoutName) = layoutKey).
            ToList()
        If Not String.IsNullOrWhiteSpace(masterName) Then
            Dim masterKey As String = NormalizePowerPointLayoutKey(masterName)
            Dim masterMatches As List(Of AutoPilotPowerPointLayoutCandidate) = matches.
                Where(Function(c) NormalizePowerPointLayoutKey(c.DesignName) = masterKey).
                ToList()
            If masterMatches.Count = 1 Then Return masterMatches(0)
            If masterMatches.Count > 1 Then Return Nothing
        End If
        If matches.Count = 1 Then Return matches(0)
        Return Nothing
    End Function

    Private Shared Function ReadAutoPilotPowerPointDesignGuidance(design As AutoPilotDesignResolution,
                                                                   templatePath As String,
                                                                   context As ToolExecutionContext) As String
        Dim guidancePath As String = ""
        If design IsNot Nothing AndAlso design.ApplicationConfig IsNot Nothing AndAlso design.Descriptor IsNot Nothing Then
            Dim relative As String = If(design.ApplicationConfig.Value(Of String)("guidance_file"), "").Trim()
            If relative <> "" Then guidancePath = design.Descriptor.ResolveRepositoryFile(relative)
        End If

        If guidancePath = "" AndAlso Not String.IsNullOrWhiteSpace(templatePath) Then
            Try
                Dim companion As String = System.IO.Path.Combine(
                    System.IO.Path.GetDirectoryName(templatePath),
                    System.IO.Path.GetFileNameWithoutExtension(templatePath) & ".md")
                If System.IO.File.Exists(companion) Then guidancePath = companion
            Catch ex As System.Exception
            End Try
        End If

        If guidancePath = "" OrElse Not System.IO.File.Exists(guidancePath) Then Return ""
        Try
            Dim text As String = System.IO.File.ReadAllText(guidancePath, System.Text.Encoding.UTF8)
            If text.Length > 24000 Then text = text.Substring(0, 24000)
            If context IsNot Nothing Then context.Log($"PowerPoint design guidance loaded: {System.IO.Path.GetFileName(guidancePath)}")
            Return text
        Catch ex As System.Exception
            If context IsNot Nothing Then context.Log($"PowerPoint design guidance could not be read: {ex.Message}")
            Return ""
        End Try
    End Function


    Private Shared Function ParsePowerPointGuidanceDirectMappings(guidance As String) As Dictionary(Of String, Tuple(Of String, String))
        Dim result As New Dictionary(Of String, Tuple(Of String, String))(StringComparer.OrdinalIgnoreCase)
        If String.IsNullOrWhiteSpace(guidance) Then Return result

        Dim inDirectSection As Boolean = False
        For Each rawLine As String In guidance.Replace(vbCrLf, vbLf).Split(ControlChars.Lf)
            Dim line As String = rawLine.Trim()
            If line.StartsWith("## ", StringComparison.Ordinal) Then
                inDirectSection = line.IndexOf("Direct mappings", StringComparison.OrdinalIgnoreCase) >= 0
                Continue For
            End If
            If Not inDirectSection OrElse Not line.StartsWith("|", StringComparison.Ordinal) Then Continue For
            Dim cells As String() = line.Trim("|"c).Split("|"c).Select(Function(v) v.Trim()).ToArray()
            If cells.Length < 2 Then Continue For
            If cells(0).IndexOf("Red Ink content type", StringComparison.OrdinalIgnoreCase) >= 0 Then Continue For
            If cells(0).Replace("-", "").Trim() = "" Then Continue For

            Dim semantic As String = cells(0).Trim().Trim("`"c).Trim()
            Dim layoutName As String = cells(1).Trim().Trim("`"c).Trim()
            Dim masterName As String = If(cells.Length >= 3, cells(2).Trim().Trim("`"c).Trim(), "")
            If semantic <> "" AndAlso layoutName <> "" Then result(semantic) = Tuple.Create(masterName, layoutName)
        Next
        Return result
    End Function

    Private Shared Function BuildPowerPointLayoutCatalogJson(catalog As IEnumerable(Of AutoPilotPowerPointLayoutCandidate)) As JArray
        Dim result As New JArray()
        If catalog Is Nothing Then Return result
        For Each c As AutoPilotPowerPointLayoutCandidate In catalog
            If c Is Nothing Then Continue For
            result.Add(New JObject From {
                {"master", c.DesignName},
                {"layout", c.LayoutName},
                {"placeholders", c.PlaceholderDetails.DeepClone()},
                {"sample_slides", c.SampleSlides.DeepClone()}
            })
        Next
        Return result
    End Function

    Private Shared Function IsPowerPointTemplateTextSlotExcludedProperty(propertyName As String) As Boolean
        Dim key As String = If(propertyName, "").Trim().ToLowerInvariant()
        Select Case key
            Case "layout", "notes", "source", "tone", "color", "data_json",
                 "template_master_name", "template_layout_name"
                Return True
        End Select
        Return False
    End Function

    Private Shared Sub CollectPowerPointTemplateTextSlots(token As JToken,
                                                           path As String,
                                                           result As JArray)
        If token Is Nothing OrElse result Is Nothing Then Return

        Select Case token.Type
            Case JTokenType.Object
                Dim obj As JObject = TryCast(token, JObject)
                If obj Is Nothing Then Return
                For Each prop As JProperty In obj.Properties()
                    If IsPowerPointTemplateTextSlotExcludedProperty(prop.Name) Then Continue For
                    Dim childPath As String = If(String.IsNullOrWhiteSpace(path), prop.Name, path & "." & prop.Name)
                    CollectPowerPointTemplateTextSlots(prop.Value, childPath, result)
                Next

            Case JTokenType.Array
                Dim arr As JArray = TryCast(token, JArray)
                If arr Is Nothing Then Return
                For i As Integer = 0 To arr.Count - 1
                    CollectPowerPointTemplateTextSlots(
                        arr(i),
                        path & "[" & (i + 1).ToString(Globalization.CultureInfo.InvariantCulture) & "]",
                        result)
                Next

            Case JTokenType.String, JTokenType.Integer, JTokenType.Float, JTokenType.Boolean
                Dim value As String = token.ToString().Trim()
                If value = "" OrElse String.IsNullOrWhiteSpace(path) Then Return
                result.Add(New JObject From {
                    {"slot_id", path},
                    {"text", value}
                })
        End Select
    End Sub

    Private Shared Function BuildPowerPointTemplateTextSlots(slideObj As JObject) As JArray
        Dim result As New JArray()
        If slideObj Is Nothing Then Return result
        CollectPowerPointTemplateTextSlots(slideObj, "", result)
        Return result
    End Function

    Private Shared Function IsAutoPilotPowerPointNativeTextEligible(slideObj As JObject,
                                                                     semanticLayout As String) As Boolean
        If slideObj Is Nothing Then Return False
        Dim layout As String = If(semanticLayout, "").Trim().ToLowerInvariant()
        If layout = "chart" OrElse layout = "table" OrElse layout = "cards" OrElse layout = "comparison" OrElse layout = "structure" OrElse layout = "process" Then Return False
        If TryCast(slideObj("chart"), JObject) IsNot Nothing Then Return False
        If TryCast(slideObj("table"), JObject) IsNot Nothing Then Return False
        If TryCast(slideObj("cards"), JArray) IsNot Nothing Then Return False
        If TryCast(slideObj("comparison"), JObject) IsNot Nothing Then Return False
        If TryCast(slideObj("structure"), JObject) IsNot Nothing Then Return False
        If TryCast(slideObj("steps"), JArray) IsNot Nothing Then Return False
        Return True
    End Function

    Private Shared Function BuildPowerPointRequestedSlidesJson(slidesArray As JArray,
                                                                unresolved As IEnumerable(Of Integer),
                                                                existingSlideCount As Integer) As JArray
        Dim result As New JArray()
        If slidesArray Is Nothing OrElse unresolved Is Nothing Then Return result
        For Each ordinal As Integer In unresolved
            If ordinal < 1 OrElse ordinal > slidesArray.Count Then Continue For
            Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
            If slideObj Is Nothing Then Continue For
            Dim semanticLayout As String =
                NormalizeAutoPilotPowerPointSemanticLayout(
                    slideObj.Value(Of String)("layout"),
                    existingSlideCount + ordinal,
                    existingSlideCount)

            result.Add(New JObject From {
                {"slide_index", ordinal},
                {"semantic_layout", semanticLayout},
                {"title", If(slideObj.Value(Of String)("title"), "")},
                {"subtitle", If(slideObj.Value(Of String)("subtitle"), "")},
                {"body", If(slideObj.Value(Of String)("body"), "")},
                {"left_title", If(slideObj.Value(Of String)("left_title"), "")},
                {"left_body", If(slideObj.Value(Of String)("left_body"), "")},
                {"right_title", If(slideObj.Value(Of String)("right_title"), "")},
                {"right_body", If(slideObj.Value(Of String)("right_body"), "")},
                {"native_text_eligible", IsAutoPilotPowerPointNativeTextEligible(slideObj, semanticLayout)},
                {"text_slots", BuildPowerPointTemplateTextSlots(slideObj)},
                {"structured_content", New JObject From {
                    {"kpis", If(TryCast(slideObj("kpis"), JArray)?.Count, 0)},
                    {"cards", If(TryCast(slideObj("cards"), JArray)?.Count, 0)},
                    {"steps", If(TryCast(slideObj("steps"), JArray)?.Count, 0)},
                    {"events", If(TryCast(slideObj("events"), JArray)?.Count, 0)},
                    {"has_table", TryCast(slideObj("table"), JObject) IsNot Nothing},
                    {"has_chart", TryCast(slideObj("chart"), JObject) IsNot Nothing},
                    {"has_structure", TryCast(slideObj("structure"), JObject) IsNot Nothing},
                    {"has_comparison", TryCast(slideObj("comparison"), JObject) IsNot Nothing},
                    {"has_matrix", TryCast(slideObj("matrix"), JObject) IsNot Nothing}
                }}
            })
        Next
        Return result
    End Function

    Private Shared Function ExtractPowerPointLayoutMappingJson(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return ""
        Dim trimmed As String = value.Trim()
        Dim firstBrace As Integer = trimmed.IndexOf("{"c)
        Dim lastBrace As Integer = trimmed.LastIndexOf("}"c)
        If firstBrace < 0 OrElse lastBrace <= firstBrace Then Return ""
        Return trimmed.Substring(firstBrace, lastBrace - firstBrace + 1)
    End Function

    Private Shared Function CloneAutoPilotPowerPointLayoutCandidate(source As AutoPilotPowerPointLayoutCandidate) As AutoPilotPowerPointLayoutCandidate
        If source Is Nothing Then Return Nothing
        Return New AutoPilotPowerPointLayoutCandidate() With {
            .DesignName = source.DesignName,
            .LayoutName = source.LayoutName,
            .SelectionReason = source.SelectionReason,
            .MasterOrdinal = source.MasterOrdinal,
            .LayoutOrdinal = source.LayoutOrdinal,
            .SlideWidth = source.SlideWidth,
            .SlideHeight = source.SlideHeight,
            .TextBindings = New JArray(),
            .PlaceholderDetails = CType(source.PlaceholderDetails.DeepClone(), JArray),
            .SampleSlides = CType(source.SampleSlides.DeepClone(), JArray)
        }
    End Function

    Private Shared Function TryBuildAutoPilotPowerPointNativeBindings(candidate As AutoPilotPowerPointLayoutCandidate,
                                                                       slideObj As JObject,
                                                                       rawBindings As JArray,
                                                                       ByRef normalizedBindings As JArray,
                                                                       ByRef errorText As String) As Boolean
        normalizedBindings = New JArray()
        errorText = ""
        If candidate Is Nothing OrElse slideObj Is Nothing Then
            errorText = "missing layout/slide metadata"
            Return False
        End If

        Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), 1, 0)
        If Not IsAutoPilotPowerPointNativeTextEligible(slideObj, semanticLayout) Then Return True

        Dim slots As JArray = BuildPowerPointTemplateTextSlots(slideObj)
        Dim slotMap As New Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
        For Each slotObj As JObject In slots.OfType(Of JObject)()
            Dim slotId As String = If(slotObj.Value(Of String)("slot_id"), "").Trim()
            Dim slotText As String = If(slotObj.Value(Of String)("text"), "")
            If slotId <> "" AndAlso slotText <> "" Then slotMap(slotId) = slotText
        Next
        If slotMap.Count = 0 Then Return True
        If rawBindings Is Nothing OrElse rawBindings.Count = 0 Then
            errorText = "no native text bindings were returned"
            Return False
        End If

        Dim placeholderMap As New Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)
        For Each ph As JObject In candidate.PlaceholderDetails.OfType(Of JObject)()
            If Not ph.Value(Of Boolean?)("text_capable").GetValueOrDefault(False) Then Continue For
            Dim key As String = If(ph.Value(Of String)("placeholder_key"), "").Trim()
            If key <> "" Then placeholderMap(key) = ph
        Next

        Dim usedSlots As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        Dim usedPlaceholders As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        For Each binding As JObject In rawBindings.OfType(Of JObject)()
            Dim placeholderKey As String = If(binding.Value(Of String)("placeholder_key"), "").Trim()
            If placeholderKey = "" OrElse Not placeholderMap.ContainsKey(placeholderKey) Then
                errorText = $"unknown/non-text placeholder_key '{placeholderKey}'"
                Return False
            End If
            If Not usedPlaceholders.Add(placeholderKey) Then
                errorText = $"placeholder_key '{placeholderKey}' was bound more than once"
                Return False
            End If

            Dim slotIds As JArray = TryCast(binding("slot_ids"), JArray)
            If slotIds Is Nothing OrElse slotIds.Count = 0 Then
                errorText = $"placeholder_key '{placeholderKey}' has no slot_ids"
                Return False
            End If

            Dim parts As New List(Of String)()
            For Each slotToken As JToken In slotIds
                Dim slotId As String = slotToken.ToString().Trim()
                If slotId = "" OrElse Not slotMap.ContainsKey(slotId) Then
                    errorText = $"unknown slot_id '{slotId}'"
                    Return False
                End If
                If Not usedSlots.Add(slotId) Then
                    errorText = $"slot_id '{slotId}' was bound more than once"
                    Return False
                End If
                parts.Add(slotMap(slotId))
            Next

            Dim ph As JObject = placeholderMap(placeholderKey)
            normalizedBindings.Add(New JObject From {
                {"placeholder_key", placeholderKey},
                {"com_type", ph.Value(Of Integer?)("com_type").GetValueOrDefault(0)},
                {"occurrence", ph.Value(Of Integer?)("occurrence").GetValueOrDefault(0)},
                {"text", String.Join(System.Environment.NewLine, parts)}
            })
        Next

        Dim missing As List(Of String) = slotMap.Keys.Where(Function(k) Not usedSlots.Contains(k)).ToList()
        If missing.Count > 0 Then
            errorText = "unbound text slot(s): " & String.Join(", ", missing)
            Return False
        End If

        Return True
    End Function

    Private Shared Function TryBuildAutoPilotPowerPointSafeFallbackBindings(candidate As AutoPilotPowerPointLayoutCandidate,
                                                                             slideObj As JObject,
                                                                             ByRef normalizedBindings As JArray,
                                                                             ByRef errorText As String) As Boolean
        normalizedBindings = New JArray()
        errorText = ""
        If candidate Is Nothing OrElse slideObj Is Nothing Then
            errorText = "missing layout/slide metadata"
            Return False
        End If

        Dim titleText As String = If(slideObj.Value(Of String)("title"), "").Trim()
        Dim contentText As String = JoinPowerPointTextParts(
            slideObj.Value(Of String)("subtitle"),
            BuildPowerPointFallbackBody(slideObj))

        Dim titlePlaceholder As JObject = Nothing
        Dim contentPlaceholder As JObject = Nothing
        Dim largestContentArea As Decimal = -1D

        For Each ph As JObject In candidate.PlaceholderDetails.OfType(Of JObject)()
            If Not ph.Value(Of Boolean?)("text_capable").GetValueOrDefault(False) Then Continue For

            Dim comType As Integer = ph.Value(Of Integer?)("com_type").GetValueOrDefault(0)
            If titlePlaceholder Is Nothing AndAlso (comType = 1 OrElse comType = 3 OrElse comType = 5) Then
                titlePlaceholder = ph
                Continue For
            End If

            If comType = 1 OrElse comType = 3 OrElse comType = 5 Then Continue For

            Dim placeholderName As String = If(ph.Value(Of String)("name"), "").Trim().ToLowerInvariant()
            If placeholderName.Contains("logo") OrElse
               placeholderName.Contains("footer") OrElse
               placeholderName.Contains("fuss") OrElse
               placeholderName.Contains("slide number") OrElse
               placeholderName.Contains("foliennummer") OrElse
               placeholderName.Contains("nummer") OrElse
               placeholderName.Contains("date") OrElse
               placeholderName.Contains("datum") Then
                Continue For
            End If

            Dim width As Decimal = ph.Value(Of Decimal?)("width").GetValueOrDefault(0D)
            Dim height As Decimal = ph.Value(Of Decimal?)("height").GetValueOrDefault(0D)
            Dim area As Decimal = width * height
            If contentPlaceholder Is Nothing OrElse area > largestContentArea Then
                contentPlaceholder = ph
                largestContentArea = area
            End If
        Next

        If titleText <> "" Then
            If titlePlaceholder Is Nothing Then
                errorText = "no safe native title placeholder is available"
                Return False
            End If
            normalizedBindings.Add(New JObject From {
                {"placeholder_key", titlePlaceholder.Value(Of String)("placeholder_key")},
                {"com_type", titlePlaceholder.Value(Of Integer?)("com_type").GetValueOrDefault(0)},
                {"occurrence", titlePlaceholder.Value(Of Integer?)("occurrence").GetValueOrDefault(0)},
                {"text", titleText}
            })
        End If

        If contentText <> "" Then
            If contentPlaceholder Is Nothing Then
                errorText = "no safe native content placeholder is available"
                Return False
            End If
            normalizedBindings.Add(New JObject From {
                {"placeholder_key", contentPlaceholder.Value(Of String)("placeholder_key")},
                {"com_type", contentPlaceholder.Value(Of Integer?)("com_type").GetValueOrDefault(0)},
                {"occurrence", contentPlaceholder.Value(Of Integer?)("occurrence").GetValueOrDefault(0)},
                {"text", contentText}
            })
        End If

        If normalizedBindings.Count = 0 Then
            errorText = "no user text required native binding"
            Return False
        End If

        Return True
    End Function

    Private Async Function ResolveAutoPilotPowerPointTemplateAssignmentsAsync(templatePath As String,
                                                                               slidesArray As JArray,
                                                                               existingSlideCount As Integer,
                                                                               design As AutoPilotDesignResolution,
                                                                               context As ToolExecutionContext,
                                                                               ct As CancellationToken) As System.Threading.Tasks.Task(Of Dictionary(Of Integer, AutoPilotPowerPointLayoutCandidate))
        Dim assignments As New Dictionary(Of Integer, AutoPilotPowerPointLayoutCandidate)()
        Dim catalog As List(Of AutoPilotPowerPointLayoutCandidate) = InspectAutoPilotPowerPointTemplateLayouts(templatePath)
        If catalog.Count = 0 OrElse slidesArray Is Nothing OrElse slidesArray.Count = 0 Then Return assignments

        Dim guidance As String = ReadAutoPilotPowerPointDesignGuidance(design, templatePath, context)
        Dim guidanceMappings As Dictionary(Of String, Tuple(Of String, String)) = ParsePowerPointGuidanceDirectMappings(guidance)
        Dim unresolved As New List(Of Integer)()
        For ordinal As Integer = 1 To slidesArray.Count
            Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
            If slideObj Is Nothing Then Continue For
            Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj.Value(Of String)("layout"), existingSlideCount + ordinal, existingSlideCount)
            Dim requestedMaster As String = If(slideObj.Value(Of String)("template_master_name"), "").Trim()
            Dim requestedLayout As String = If(slideObj.Value(Of String)("template_layout_name"), "").Trim()
            Dim configuredMaster As String = ""
            Dim configuredLayout As String = ""
            ResolveConfiguredPowerPointLayoutTarget(design, semanticLayout, configuredMaster, configuredLayout)

            Dim guidedMaster As String = ""
            Dim guidedLayout As String = ""
            If guidanceMappings.ContainsKey(semanticLayout) Then
                Dim directTarget As Tuple(Of String, String) = guidanceMappings(semanticLayout)
                guidedMaster = directTarget.Item1
                guidedLayout = directTarget.Item2
            End If

            ' Binding precedence is explicit per-slide selection, then human-readable Markdown
            ' design policy, then the legacy JSON layout_map. Markdown is intentionally ahead
            ' of JSON because it is the user-friendly design-owner contract.
            Dim exactMaster As String = requestedMaster
            Dim exactLayout As String = requestedLayout
            Dim selectionReason As String = "explicit per-slide layout"
            If exactLayout = "" AndAlso guidedLayout <> "" Then
                exactMaster = guidedMaster
                exactLayout = guidedLayout
                selectionReason = "explicit Markdown direct mapping"
            ElseIf exactLayout = "" AndAlso configuredLayout <> "" Then
                exactMaster = configuredMaster
                exactLayout = configuredLayout
                selectionReason = "configured layout_map"
            End If

            If exactLayout <> "" Then
                Dim exactMatch As AutoPilotPowerPointLayoutCandidate = FindPowerPointCatalogLayout(catalog, exactMaster, exactLayout)
                If exactMatch IsNot Nothing Then
                    Dim exactAssignment As AutoPilotPowerPointLayoutCandidate = CloneAutoPilotPowerPointLayoutCandidate(exactMatch)
                    exactAssignment.SelectionReason = selectionReason
                    exactAssignment = ResolveAutoPilotPowerPointVisualCanvasLayout(catalog, exactAssignment, semanticLayout, selectionReason, context, ordinal)
                    assignments(ordinal) = exactAssignment
                    If context IsNot Nothing AndAlso selectionReason = "explicit Markdown direct mapping" Then
                        context.Log($"PowerPoint Markdown direct mapping resolved locally: slide={ordinal}; semantic='{semanticLayout}'; master='{exactAssignment.DesignName}'; layout='{exactAssignment.LayoutName}'.", "diag")
                    End If
                    Continue For
                End If
                If context IsNot Nothing Then context.Log($"PowerPoint exact layout mapping not found: slide={ordinal}; master='{exactMaster}'; layout='{exactLayout}'. Falling back to template interpretation.")
            End If
            unresolved.Add(ordinal)
        Next

        If unresolved.Count = 0 Then Return assignments

        Dim payload As New JObject From {
            {"guidance", guidance},
            {"layouts", BuildPowerPointLayoutCatalogJson(catalog)},
            {"slides", BuildPowerPointRequestedSlidesJson(slidesArray, unresolved, existingSlideCount)}
        }

        Dim systemPrompt As String =
            "You are a PowerPoint template layout and placeholder mapper. Map each requested slide to an exact master/layout from the supplied template metadata. " &
            "Use only the supplied human guidance, actual master/layout names, placeholder metadata/geometry, and sample-slide content. " &
            "Do not use numeric layout positions, brand knowledge, fixed layout-name conventions, or hard-coded assumptions about how a template should be organized. " &
            "Human guidance is authoritative when it is explicit. When guidance is absent or incomplete, infer cautiously from the actual template structure and sample slides. " &
            "A requested semantic_layout describes the intended content, not a required template layout name. " &
            "For slides where native_text_eligible=true, also map EVERY supplied text_slots entry exactly once to a text-capable placeholder of the chosen layout. " &
            "Use only exact placeholder_key values supplied by that layout; a numeric suffix inside placeholder_key is an opaque host identity, not a layout-position hint. Multiple slot_ids may be assigned to one placeholder when they belong together; do not invent, rewrite, omit, or duplicate text. " &
            "Never bind content to footer, date, slide-number, picture, chart, table, media, or other non-text placeholders, and do not overwrite a placeholder whose supplied name/sample_text shows that it is fixed branding, a logo, or other decorative template content. " &
            "If no template layout can safely accommodate the complete slide, return an empty master/layout rather than guessing. " &
            "Return ONLY JSON in the form {""assignments"":[{""slide_index"":1,""master"":""exact master name"",""layout"":""exact layout name"",""reason"":""brief reason"",""text_bindings"":[{""placeholder_key"":""body:0"",""slot_ids"":[""body""]}]}]} and include every supplied slide_index exactly once."

        Try
            If context IsNot Nothing Then context.Log($"PowerPoint template interpretation started: layouts={catalog.Count}; slides={unresolved.Count}; guidance={If(guidance = "", "none", "present")}")
            Dim templatePayload As String = payload.ToString(Newtonsoft.Json.Formatting.None)
            Dim configuredTemplateTimeoutMs As Long = 0L
            If context IsNot Nothing AndAlso context.ToolingModel IsNot Nothing AndAlso context.ToolingModel.Timeout > 0 Then
                configuredTemplateTimeoutMs = context.ToolingModel.Timeout
            ElseIf INI_Timeout > 0 Then
                configuredTemplateTimeoutMs = INI_Timeout
            Else
                configuredTemplateTimeoutMs = 30000L
            End If
            Dim templateCallTimeoutMs As Integer = SharedLibrary.Agents.HostToolRegistration.GetPerCallLlmTimeoutMs(
                configuredTemplateTimeoutMs,
                New String() {"create_powerpoint"},
                systemPrompt.Length,
                templatePayload.Length)
            If context IsNot Nothing AndAlso templateCallTimeoutMs <> configuredTemplateTimeoutMs Then
                context.Log($"PowerPoint template interpretation timeout elevated per heavy-call policy: baseMs={configuredTemplateTimeoutMs}; effectiveMs={templateCallTimeoutMs}.", "diag")
            End If
            Dim raw As String = Await LLM(systemPrompt,
                                          templatePayload,
                                          Temperature:="0",
                                          Timeout:=templateCallTimeoutMs,
                                          HideSplash:=True,
                                          cancellationToken:=ct,
                                          EnsureUI:=False,
                                          ToolExecution:=True).ConfigureAwait(False)
            Dim json As String = ExtractPowerPointLayoutMappingJson(raw)
            If json = "" Then Return assignments
            Dim root As JObject = JObject.Parse(json)
            Dim resultArray As JArray = TryCast(root("assignments"), JArray)
            If resultArray Is Nothing Then Return assignments
            For Each item As JObject In resultArray.OfType(Of JObject)()
                Dim ordinal As Integer = item.Value(Of Integer?)("slide_index").GetValueOrDefault(0)
                If Not unresolved.Contains(ordinal) Then Continue For
                Dim masterName As String = If(item.Value(Of String)("master"), "").Trim()
                Dim layoutName As String = If(item.Value(Of String)("layout"), "").Trim()
                If layoutName = "" Then Continue For
                Dim catalogCandidate As AutoPilotPowerPointLayoutCandidate = FindPowerPointCatalogLayout(catalog, masterName, layoutName)
                If catalogCandidate Is Nothing Then
                    If context IsNot Nothing Then context.Log($"PowerPoint template interpretation returned an unknown/ambiguous layout: slide={ordinal}; master='{masterName}'; layout='{layoutName}'.")
                    Continue For
                End If

                Dim candidate As AutoPilotPowerPointLayoutCandidate = CloneAutoPilotPowerPointLayoutCandidate(catalogCandidate)
                candidate.SelectionReason = If(item.Value(Of String)("reason"), "LLM template interpretation")

                Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
                Dim semanticLayout As String = NormalizeAutoPilotPowerPointSemanticLayout(slideObj?.Value(Of String)("layout"), existingSlideCount + ordinal, existingSlideCount)
                candidate = ResolveAutoPilotPowerPointVisualCanvasLayout(catalog, candidate, semanticLayout, candidate.SelectionReason, context, ordinal)
                If IsAutoPilotPowerPointNativeTextEligible(slideObj, semanticLayout) Then
                    Dim normalizedBindings As JArray = Nothing
                    Dim bindingError As String = ""
                    If Not TryBuildAutoPilotPowerPointNativeBindings(
                        candidate,
                        slideObj,
                        TryCast(item("text_bindings"), JArray),
                        normalizedBindings,
                        bindingError) Then

                        If context IsNot Nothing Then context.Log($"PowerPoint native placeholder binding rejected: slide={ordinal}; layout='{layoutName}'; reason={bindingError}")

                        Dim fallbackBindings As JArray = Nothing
                        Dim fallbackBindingError As String = ""
                        If Not TryBuildAutoPilotPowerPointSafeFallbackBindings(
                            candidate,
                            slideObj,
                            fallbackBindings,
                            fallbackBindingError) Then

                            If context IsNot Nothing Then context.Log($"PowerPoint safe native fallback binding rejected: slide={ordinal}; layout='{layoutName}'; reason={fallbackBindingError}")

                            Dim repairedCandidate As AutoPilotPowerPointLayoutCandidate = Nothing
                            Dim repairedBindings As JArray = Nothing
                            For Each alternativeCatalogCandidate As AutoPilotPowerPointLayoutCandidate In catalog
                                If alternativeCatalogCandidate Is Nothing Then Continue For
                                If Not String.IsNullOrWhiteSpace(masterName) AndAlso
                                   Not String.Equals(alternativeCatalogCandidate.DesignName, masterName, StringComparison.OrdinalIgnoreCase) Then Continue For

                                Dim alternativeCandidate As AutoPilotPowerPointLayoutCandidate = CloneAutoPilotPowerPointLayoutCandidate(alternativeCatalogCandidate)
                                Dim alternativeBindings As JArray = Nothing
                                Dim alternativeError As String = ""
                                If TryBuildAutoPilotPowerPointSafeFallbackBindings(alternativeCandidate, slideObj, alternativeBindings, alternativeError) Then
                                    repairedCandidate = alternativeCandidate
                                    repairedBindings = alternativeBindings
                                    Exit For
                                End If
                            Next

                            If repairedCandidate Is Nothing Then
                                For Each alternativeCatalogCandidate As AutoPilotPowerPointLayoutCandidate In catalog
                                    If alternativeCatalogCandidate Is Nothing Then Continue For
                                    Dim alternativeCandidate As AutoPilotPowerPointLayoutCandidate = CloneAutoPilotPowerPointLayoutCandidate(alternativeCatalogCandidate)
                                    Dim alternativeBindings As JArray = Nothing
                                    Dim alternativeError As String = ""
                                    If TryBuildAutoPilotPowerPointSafeFallbackBindings(alternativeCandidate, slideObj, alternativeBindings, alternativeError) Then
                                        repairedCandidate = alternativeCandidate
                                        repairedBindings = alternativeBindings
                                        Exit For
                                    End If
                                Next
                            End If

                            If repairedCandidate Is Nothing Then Continue For
                            repairedCandidate.SelectionReason = $"safe native layout repair after unsuitable interpreted layout '{layoutName}'"
                            repairedCandidate.TextBindings = repairedBindings
                            candidate = repairedCandidate
                            normalizedBindings = repairedBindings
                            If context IsNot Nothing Then context.Log($"PowerPoint safe native layout repair applied: slide={ordinal}; rejectedLayout='{layoutName}'; repairedLayout='{candidate.LayoutName}'.", "diag")
                        Else
                            normalizedBindings = fallbackBindings
                            If context IsNot Nothing Then context.Log($"PowerPoint safe native fallback binding applied: slide={ordinal}; layout='{layoutName}'.", "diag")
                        End If

                        If normalizedBindings Is Nothing Then normalizedBindings = fallbackBindings
                    End If
                    candidate.TextBindings = normalizedBindings
                End If

                assignments(ordinal) = candidate
            Next
            If context IsNot Nothing Then context.Log($"PowerPoint template interpretation completed: assigned={assignments.Count}/{slidesArray.Count}.")
        Catch ex As System.Exception
            If context IsNot Nothing Then context.Log($"PowerPoint template interpretation failed; native template binding remains blocked. {ex.Message}")
        End Try

        Return assignments
    End Function

    Private Shared Function IsAutoPilotPowerPointRichVisualSemantic(semanticLayout As String) As Boolean
        Select Case If(semanticLayout, "").Trim().ToLowerInvariant()
            Case "cards", "comparison", "structure", "process", "timeline", "matrix", "chart", "kpi"
                Return True
        End Select
        Return False
    End Function

    Private Shared Function IsAutoPilotPowerPointVisualCanvasPlaceholderType(comType As Integer) As Boolean
        Return comType = 2 OrElse
               comType = 6 OrElse
               comType = 7 OrElse
               comType = 8 OrElse
               comType = 9 OrElse
               comType = 11 OrElse
               comType = 12 OrElse
               comType = 17 OrElse
               comType = 18
    End Function

    Private Shared Function ScoreAutoPilotPowerPointVisualCanvasCandidate(candidate As AutoPilotPowerPointLayoutCandidate,
                                                                            semanticLayout As String,
                                                                            ByRef usableZoneCount As Integer) As Double
        usableZoneCount = 0
        If candidate Is Nothing OrElse candidate.PlaceholderDetails Is Nothing Then Return -1.0R

        Dim hasTitle As Boolean = False
        Dim generalAreas As New List(Of Double)()
        For Each placeholderObj As JObject In candidate.PlaceholderDetails.OfType(Of JObject)()
            Dim comType As Integer = placeholderObj.Value(Of Integer?)("com_type").GetValueOrDefault(0)
            If comType = 1 OrElse comType = 3 OrElse comType = 5 Then hasTitle = True

            ' A neutral visual canvas should be based on general-purpose content/object placeholders,
            ' not on a large picture/chart/table placeholder whose geometry happens to be bigger.
            ' This keeps the resolver design-agnostic while avoiding template-specific special regions.
            If Not (comType = 2 OrElse comType = 6 OrElse comType = 7 OrElse comType = 17) Then Continue For

            Dim width As Double = CDbl(placeholderObj.Value(Of Decimal?)("width").GetValueOrDefault(0D))
            Dim height As Double = CDbl(placeholderObj.Value(Of Decimal?)("height").GetValueOrDefault(0D))
            If width <= 0.0R OrElse height <= 0.0R Then Continue For
            If candidate.SlideWidth > 0L AndAlso width < CDbl(candidate.SlideWidth) * 0.18R Then Continue For
            If candidate.SlideHeight > 0L AndAlso height < CDbl(candidate.SlideHeight) * 0.20R Then Continue For
            generalAreas.Add(width * height)
        Next

        If Not hasTitle OrElse generalAreas.Count = 0 Then Return -1.0R
        generalAreas.Sort()
        generalAreas.Reverse()
        usableZoneCount = generalAreas.Count

        Dim semantic As String = If(semanticLayout, "").Trim().ToLowerInvariant()
        If semantic = "comparison" AndAlso generalAreas.Count >= 2 Then
            Dim firstArea As Double = generalAreas(0)
            Dim secondArea As Double = generalAreas(1)
            Dim balance As Double = If(firstArea > 0.0R, secondArea / firstArea, 0.0R)
            Return (firstArea + secondArea) * Math.Max(0.25R, balance)
        End If

        Dim score As Double = generalAreas(0)
        For i As Integer = 1 To generalAreas.Count - 1
            score += generalAreas(i) * 0.08R
        Next
        If semantic = "comparison" Then score *= 0.50R
        If semantic <> "comparison" AndAlso generalAreas.Count = 1 Then score *= 1.10R
        Return score
    End Function

    Private Shared Function FindBestAutoPilotPowerPointVisualCanvasLayout(catalog As IEnumerable(Of AutoPilotPowerPointLayoutCandidate),
                                                                            preferredMaster As String,
                                                                            semanticLayout As String) As AutoPilotPowerPointLayoutCandidate
        If catalog Is Nothing Then Return Nothing
        Dim best As AutoPilotPowerPointLayoutCandidate = Nothing
        Dim bestScore As Double = -1.0R
        Dim preferredMasterKey As String = NormalizePowerPointLayoutKey(preferredMaster)

        For Each candidate As AutoPilotPowerPointLayoutCandidate In catalog
            If candidate Is Nothing Then Continue For
            Dim zoneCount As Integer = 0
            Dim score As Double = ScoreAutoPilotPowerPointVisualCanvasCandidate(candidate, semanticLayout, zoneCount)
            If score < 0.0R Then Continue For

            If preferredMasterKey <> "" AndAlso NormalizePowerPointLayoutKey(candidate.DesignName) = preferredMasterKey Then
                score *= 1.08R
            End If

            If score > bestScore Then
                best = candidate
                bestScore = score
            End If
        Next
        Return best
    End Function

    Private Shared Function ResolveAutoPilotPowerPointVisualCanvasLayout(catalog As IEnumerable(Of AutoPilotPowerPointLayoutCandidate),
                                                                           mappedCandidate As AutoPilotPowerPointLayoutCandidate,
                                                                           semanticLayout As String,
                                                                           selectionReason As String,
                                                                           context As ToolExecutionContext,
                                                                           slideOrdinal As Integer) As AutoPilotPowerPointLayoutCandidate
        If mappedCandidate Is Nothing OrElse Not IsAutoPilotPowerPointRichVisualSemantic(semanticLayout) Then Return mappedCandidate

        Dim mappedZoneCount As Integer = 0
        Dim mappedScore As Double = ScoreAutoPilotPowerPointVisualCanvasCandidate(mappedCandidate, semanticLayout, mappedZoneCount)
        If mappedScore >= 0.0R Then Return mappedCandidate

        Dim fallback As AutoPilotPowerPointLayoutCandidate = FindBestAutoPilotPowerPointVisualCanvasLayout(catalog, mappedCandidate.DesignName, semanticLayout)
        If fallback Is Nothing Then Return mappedCandidate

        Dim repaired As AutoPilotPowerPointLayoutCandidate = CloneAutoPilotPowerPointLayoutCandidate(fallback)
        repaired.SelectionReason = selectionReason & "; geometry-validated visual canvas fallback"
        If context IsNot Nothing Then
            context.Log($"PowerPoint visual canvas resolver: slide={slideOrdinal}; semantic='{semanticLayout}'; rejected layout='{mappedCandidate.LayoutName}' because it exposes no usable rich-content zone; selected master='{repaired.DesignName}'; layout='{repaired.LayoutName}'.", "diag")
        End If
        Return repaired
    End Function

    Private Shared Function AddAutoPilotPowerPointSlide(pres As Object,
                                                         slideIndex As Integer,
                                                         blankLayoutId As Integer,
                                                         templateLayout As AutoPilotPowerPointLayoutCandidate,
                                                         ByRef templateLayoutApplied As Boolean) As Object
        templateLayoutApplied = False
        If pres Is Nothing Then Return Nothing
        If templateLayout Is Nothing Then Return pres.Slides.Add(slideIndex, blankLayoutId)

        Dim designs As Object = Nothing
        Dim exactMasterMatches As New List(Of Object)()
        Dim uniqueLayoutMatches As New List(Of Object)()
        Dim ordinalLayout As Object = Nothing
        Try
            designs = pres.Designs
            Dim wantedLayoutKey As String = NormalizePowerPointLayoutKey(templateLayout.LayoutName)
            Dim wantedMasterKey As String = NormalizePowerPointLayoutKey(templateLayout.DesignName)
            Dim syntheticMaster As Boolean = IsSyntheticPowerPointMasterName(templateLayout.DesignName)

            Dim designCount As Integer = CInt(designs.Count)
            For designIndex As Integer = 1 To designCount
                Dim designObject As Object = Nothing
                Dim masterObject As Object = Nothing
                Dim layouts As Object = Nothing
                Try
                    designObject = designs(designIndex)
                    masterObject = designObject.SlideMaster
                    layouts = masterObject.CustomLayouts

                    Dim liveMasterName As String = ""
                    Try : liveMasterName = CStr(designObject.Name) : Catch ex As System.Exception : End Try
                    If liveMasterName = "" Then Try : liveMasterName = CStr(masterObject.Name) : Catch ex As System.Exception : End Try
                    Dim masterMatches As Boolean = syntheticMaster OrElse
                        NormalizePowerPointLayoutKey(liveMasterName) = wantedMasterKey

                    Dim layoutCount As Integer = CInt(layouts.Count)
                    For layoutIndex As Integer = 1 To layoutCount
                        Dim layoutObject As Object = Nothing
                        Dim keepLayoutObject As Boolean = False
                        Try
                            layoutObject = layouts(layoutIndex)
                            Dim liveLayoutName As String = ""
                            Try : liveLayoutName = CStr(layoutObject.Name) : Catch ex As System.Exception : End Try
                            If NormalizePowerPointLayoutKey(liveLayoutName) <> wantedLayoutKey Then Continue For

                            uniqueLayoutMatches.Add(layoutObject)
                            If masterMatches Then exactMasterMatches.Add(layoutObject)
                            keepLayoutObject = True
                        Finally
                            If layoutObject IsNot Nothing AndAlso Not keepLayoutObject Then
                                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layoutObject) : Catch ex As System.Exception : End Try
                            End If
                        End Try
                    Next
                Finally
                    If layouts IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layouts) : Catch ex As System.Exception : End Try
                    If masterObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(masterObject) : Catch ex As System.Exception : End Try
                    If designObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(designObject) : Catch ex As System.Exception : End Try
                End Try
            Next

            Dim selectedLayout As Object = Nothing
            If exactMasterMatches.Count = 1 Then
                selectedLayout = exactMasterMatches(0)
            ElseIf uniqueLayoutMatches.Count = 1 Then
                selectedLayout = uniqueLayoutMatches(0)
            End If

            ' The Open XML catalog and the live PowerPoint presentation are the same
            ' template package. If human-readable names are unavailable/ambiguous,
            ' use the exact master/layout order recorded from p:sldMasterIdLst and
            ' p:sldLayoutIdLst. This is structural identity, not semantic guessing.
            If selectedLayout Is Nothing AndAlso
               templateLayout.MasterOrdinal > 0 AndAlso
               templateLayout.LayoutOrdinal > 0 AndAlso
               templateLayout.MasterOrdinal <= designCount Then

                Dim designObject As Object = Nothing
                Dim masterObject As Object = Nothing
                Dim layouts As Object = Nothing
                Try
                    designObject = designs(templateLayout.MasterOrdinal)
                    masterObject = designObject.SlideMaster
                    layouts = masterObject.CustomLayouts
                    If templateLayout.LayoutOrdinal <= CInt(layouts.Count) Then
                        ordinalLayout = layouts(templateLayout.LayoutOrdinal)
                        Dim liveLayoutName As String = ""
                        Try : liveLayoutName = CStr(ordinalLayout.Name) : Catch ex As System.Exception : End Try

                        ' When the package carries a real layout name, require the live
                        ' COM object at that structural position to report the same name.
                        ' This prevents an ordinal mismatch from silently selecting the
                        ' wrong master/layout in unusual multi-master presentations.
                        If String.IsNullOrWhiteSpace(templateLayout.LayoutName) OrElse
                           System.Text.RegularExpressions.Regex.IsMatch(templateLayout.LayoutName.Trim(), "^Layout\s+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) OrElse
                           NormalizePowerPointLayoutKey(liveLayoutName) = wantedLayoutKey Then

                            selectedLayout = ordinalLayout
                        End If
                    End If
                Finally
                    If layouts IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layouts) : Catch ex As System.Exception : End Try
                    If masterObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(masterObject) : Catch ex As System.Exception : End Try
                    If designObject IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(designObject) : Catch ex As System.Exception : End Try
                End Try
            End If

            If selectedLayout Is Nothing Then
                Return pres.Slides.Add(slideIndex, blankLayoutId)
            End If

            Dim createdSlide As Object = Nothing
            Try
                createdSlide = pres.Slides.AddSlide(slideIndex, selectedLayout)
                Dim actualLayout As Object = Nothing
                Try
                    actualLayout = createdSlide.CustomLayout
                    Dim actualLayoutName As String = ""
                    Try : actualLayoutName = CStr(actualLayout.Name) : Catch ex As System.Exception : End Try
                    If Not String.IsNullOrWhiteSpace(templateLayout.LayoutName) AndAlso
                       Not System.Text.RegularExpressions.Regex.IsMatch(templateLayout.LayoutName.Trim(), "^Layout\s+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) AndAlso
                       NormalizePowerPointLayoutKey(actualLayoutName) <> wantedLayoutKey Then

                        Try : createdSlide.Delete() : Catch ex As System.Exception : End Try
                        Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(createdSlide) : Catch ex As System.Exception : End Try
                        createdSlide = Nothing
                        Return pres.Slides.Add(slideIndex, blankLayoutId)
                    End If
                Finally
                    If actualLayout IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(actualLayout) : Catch ex As System.Exception : End Try
                End Try

                templateLayoutApplied = True
                Return createdSlide
            Catch ex As System.Exception
                If createdSlide IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(createdSlide) : Catch inner As System.Exception : End Try
                End If
                Return pres.Slides.Add(slideIndex, blankLayoutId)
            End Try
        Catch ex As System.Exception
            Return pres.Slides.Add(slideIndex, blankLayoutId)
        Finally
            Dim released As New HashSet(Of Integer)()
            For Each layoutObject As Object In uniqueLayoutMatches
                If layoutObject Is Nothing Then Continue For
                Try
                    Dim key As Integer = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(layoutObject)
                    If released.Add(key) Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(layoutObject)
                Catch ex As System.Exception
                End Try
            Next
            If ordinalLayout IsNot Nothing Then
                Try
                    Dim key As Integer = System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(ordinalLayout)
                    If released.Add(key) Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(ordinalLayout)
                Catch ex As System.Exception
                End Try
            End If
            If designs IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(designs) : Catch ex As System.Exception : End Try
        End Try
    End Function

    Private Shared Function TrySetPowerPointPlaceholderText(sld As Object,
                                                             text As String,
                                                             preferredTypes As IEnumerable(Of Integer),
                                                             occurrence As Integer) As Boolean
        If sld Is Nothing OrElse String.IsNullOrWhiteSpace(text) OrElse preferredTypes Is Nothing Then Return False
        If occurrence < 0 Then occurrence = 0

        Dim shapes As Object = Nothing
        Try
            shapes = sld.Shapes
            For Each preferredType As Integer In preferredTypes
                Dim matched As Integer = 0
                Dim shapeCount As Integer = CInt(shapes.Count)
                For shapeIndex As Integer = 1 To shapeCount
                    Dim shape As Object = Nothing
                    Try
                        shape = shapes(shapeIndex)
                        Dim placeholderType As Integer
                        Try
                            placeholderType = CInt(shape.PlaceholderFormat.Type)
                        Catch ex As System.Exception
                            Continue For
                        End Try

                        If placeholderType <> preferredType Then Continue For
                        If matched = occurrence Then
                            shape.TextFrame.TextRange.Text = text
                            Return True
                        End If
                        matched += 1
                    Catch ex As System.Exception
                    Finally
                        If shape IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch ex As System.Exception : End Try
                        End If
                    End Try
                Next
            Next
        Catch ex As System.Exception
        Finally
            If shapes IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch ex As System.Exception : End Try
            End If
        End Try
        Return False
    End Function

    Private Shared Function ApplyAutoPilotPowerPointBulletItemsToShape(shape As Object,
                                                                        bulletItems As JArray) As Boolean
        If shape Is Nothing OrElse bulletItems Is Nothing OrElse bulletItems.Count = 0 Then Return False

        Dim lines As New System.Collections.Generic.List(Of String)()
        Dim levels As New System.Collections.Generic.List(Of Integer)()
        For Each itemObj As JObject In bulletItems.OfType(Of JObject)()
            Dim itemText As String = If(itemObj.Value(Of String)("text"), "").Trim()
            If itemText = "" Then Continue For
            lines.Add(itemText)
            levels.Add(Math.Max(0, Math.Min(4, itemObj.Value(Of Integer?)("level").GetValueOrDefault(0))))
        Next
        If lines.Count = 0 Then Return False

        Try
            shape.TextFrame.TextRange.Text = String.Join(vbCrLf, lines)
            Dim baseFontSize As Single = 18.0F
            Try
                baseFontSize = CSng(shape.TextFrame.TextRange.Font.Size)
                If baseFontSize <= 0.0F OrElse baseFontSize > 60.0F Then baseFontSize = 18.0F
            Catch ex As System.Exception
                baseFontSize = 18.0F
            End Try

            For i As Integer = 1 To lines.Count
                Dim paragraph As Object = Nothing
                Try
                    paragraph = shape.TextFrame.TextRange.Paragraphs(i, 1)
                    Dim level As Integer = levels(i - 1)
                    paragraph.ParagraphFormat.Bullet.Visible = -1
                    Try : paragraph.ParagraphFormat.Bullet.Type = 1 : Catch ex As System.Exception : End Try
                    paragraph.IndentLevel = Math.Min(5, level + 1)
                    ' Keep the template's native typography for every indentation level.
                    ' IndentLevel selects the design carrier's lvl1/lvl2/... paragraph style;
                    ' the renderer must not impose a cross-template font-size scale here.
                    Try
                        paragraph.ParagraphFormat.SpaceAfter = If(level = 0, 9.0F, 4.0F)
                    Catch ex As System.Exception
                    End Try
                Finally
                    If paragraph IsNot Nothing Then
                        Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(paragraph) : Catch ex As System.Exception : End Try
                    End If
                End Try
            Next
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function

    Private Shared Function TrySetPowerPointPlaceholderBulletItems(sld As Object,
                                                                    bulletItems As JArray,
                                                                    preferredTypes As System.Collections.Generic.IEnumerable(Of Integer),
                                                                    occurrence As Integer) As Boolean
        If sld Is Nothing OrElse bulletItems Is Nothing OrElse bulletItems.Count = 0 OrElse preferredTypes Is Nothing Then Return False
        If occurrence < 0 Then occurrence = 0

        Dim shapes As Object = Nothing
        Try
            shapes = sld.Shapes
            For Each preferredType As Integer In preferredTypes
                Dim matched As Integer = 0
                Dim shapeCount As Integer = CInt(shapes.Count)
                For shapeIndex As Integer = 1 To shapeCount
                    Dim shape As Object = Nothing
                    Try
                        shape = shapes(shapeIndex)
                        Dim placeholderType As Integer
                        Try
                            placeholderType = CInt(shape.PlaceholderFormat.Type)
                        Catch ex As System.Exception
                            Continue For
                        End Try
                        If placeholderType <> preferredType Then Continue For
                        If matched = occurrence Then
                            Return ApplyAutoPilotPowerPointBulletItemsToShape(shape, bulletItems)
                        End If
                        matched += 1
                    Finally
                        If shape IsNot Nothing Then
                            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch ex As System.Exception : End Try
                        End If
                    End Try
                Next
            Next
        Catch ex As System.Exception
        Finally
            If shapes IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch ex As System.Exception : End Try
            End If
        End Try
        Return False
    End Function

    Private Shared Function RenderAutoPilotPowerPointNativeTextBindings(sld As Object,
                                                                        templateLayout As AutoPilotPowerPointLayoutCandidate) As Boolean
        If sld Is Nothing OrElse templateLayout Is Nothing OrElse templateLayout.TextBindings Is Nothing OrElse templateLayout.TextBindings.Count = 0 Then
            Return False
        End If

        For Each binding As JObject In templateLayout.TextBindings.OfType(Of JObject)()
            Dim comType As Integer = binding.Value(Of Integer?)("com_type").GetValueOrDefault(0)
            Dim occurrence As Integer = binding.Value(Of Integer?)("occurrence").GetValueOrDefault(0)
            Dim value As String = If(binding.Value(Of String)("text"), "")
            If comType <= 0 OrElse String.IsNullOrWhiteSpace(value) Then Return False
            Dim applied As Boolean
            If comType = 2 OrElse comType = 6 OrElse comType = 7 OrElse comType = 17 Then
                applied = TrySetPowerPointBestNarrativePlaceholderText(sld, value, New Integer() {2, 7, 17, 6}, occurrence)
            Else
                applied = TrySetPowerPointPlaceholderText(sld, value, New Integer() {comType}, occurrence)
            End If
            If Not applied Then Return False
        Next

        Return True
    End Function

    Private Shared Function TrySetPowerPointTemplateTitle(sld As Object, title As String) As Boolean
        Return TrySetPowerPointPlaceholderText(
            sld,
            title,
            New Integer() {1, 3, 5},
            0
        )
    End Function

    Private Shared Function GetPowerPointNarrativePlaceholderIndices(sld As Object,
                                                                        preferredTypes As System.Collections.Generic.IEnumerable(Of Integer)) As System.Collections.Generic.List(Of Integer)
        Dim result As New System.Collections.Generic.List(Of Integer)()
        If sld Is Nothing OrElse preferredTypes Is Nothing Then Return result

        Dim preferred As New System.Collections.Generic.HashSet(Of Integer)(preferredTypes)
        Dim candidates As New System.Collections.Generic.List(Of System.Tuple(Of Integer, Double, Double, Double))()
        Dim shapes As Object = Nothing
        Try
            shapes = sld.Shapes
            Dim shapeCount As Integer = CInt(shapes.Count)
            For shapeIndex As Integer = 1 To shapeCount
                Dim shape As Object = Nothing
                Try
                    shape = shapes(shapeIndex)
                    Dim placeholderType As Integer
                    Try
                        placeholderType = CInt(shape.PlaceholderFormat.Type)
                    Catch ex As System.Exception
                        Continue For
                    End Try
                    If Not preferred.Contains(placeholderType) Then Continue For

                    Dim width As Double = 0.0R
                    Dim height As Double = 0.0R
                    Dim left As Double = 0.0R
                    Dim top As Double = 0.0R
                    Try
                        width = CDbl(shape.Width)
                        height = CDbl(shape.Height)
                        left = CDbl(shape.Left)
                        top = CDbl(shape.Top)
                    Catch ex As System.Exception
                        Continue For
                    End Try
                    If width <= 0.0R OrElse height <= 0.0R Then Continue For
                    candidates.Add(System.Tuple.Create(shapeIndex, width * height, left, top))
                Finally
                    If shape IsNot Nothing Then
                        Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch ex As System.Exception : End Try
                    End If
                End Try
            Next
        Catch ex As System.Exception
        Finally
            If shapes IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch ex As System.Exception : End Try
            End If
        End Try

        If candidates.Count = 0 Then Return result

        ' Placeholder COM types are not semantically reliable across customer templates.
        ' Keep only genuinely usable narrative regions relative to the largest candidate,
        ' then order spatially so two-column occurrence 0/1 remains left-to-right.
        Dim maxArea As Double = candidates.Max(Function(candidate) candidate.Item2)
        Dim usable As New System.Collections.Generic.List(Of System.Tuple(Of Integer, Double, Double, Double))()
        For Each candidate In candidates
            If candidate.Item2 >= maxArea * 0.3R Then usable.Add(candidate)
        Next
        usable.Sort(
            Function(first, second)
                Dim leftComparison As Integer = first.Item3.CompareTo(second.Item3)
                If leftComparison <> 0 Then Return leftComparison
                Dim topComparison As Integer = first.Item4.CompareTo(second.Item4)
                If topComparison <> 0 Then Return topComparison
                Return second.Item2.CompareTo(first.Item2)
            End Function)

        For Each candidate In usable
            result.Add(candidate.Item1)
        Next
        Return result
    End Function

    Private Shared Function TrySetPowerPointBestNarrativePlaceholderText(sld As Object,
                                                                           text As String,
                                                                           preferredTypes As System.Collections.Generic.IEnumerable(Of Integer),
                                                                           occurrence As Integer) As Boolean
        If sld Is Nothing OrElse String.IsNullOrWhiteSpace(text) Then Return False
        If occurrence < 0 Then occurrence = 0
        Dim indices As System.Collections.Generic.List(Of Integer) = GetPowerPointNarrativePlaceholderIndices(sld, preferredTypes)
        If occurrence >= indices.Count Then Return False

        Dim shape As Object = Nothing
        Try
            shape = sld.Shapes(indices(occurrence))
            shape.TextFrame.TextRange.Text = text
            Return True
        Catch ex As System.Exception
            Return False
        Finally
            If shape IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch ex As System.Exception : End Try
            End If
        End Try
    End Function

    Private Shared Function TrySetPowerPointBestNarrativePlaceholderBulletItems(sld As Object,
                                                                                  bulletItems As JArray,
                                                                                  preferredTypes As System.Collections.Generic.IEnumerable(Of Integer),
                                                                                  occurrence As Integer) As Boolean
        If sld Is Nothing OrElse bulletItems Is Nothing OrElse bulletItems.Count = 0 Then Return False
        If occurrence < 0 Then occurrence = 0
        Dim indices As System.Collections.Generic.List(Of Integer) = GetPowerPointNarrativePlaceholderIndices(sld, preferredTypes)
        If occurrence >= indices.Count Then Return False

        Dim shape As Object = Nothing
        Try
            shape = sld.Shapes(indices(occurrence))
            Return ApplyAutoPilotPowerPointBulletItemsToShape(shape, bulletItems)
        Catch ex As System.Exception
            Return False
        Finally
            If shape IsNot Nothing Then
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch ex As System.Exception : End Try
            End If
        End Try
    End Function

    Private Shared Function TrySetPowerPointTemplateContent(sld As Object,
                                                             text As String,
                                                             occurrence As Integer) As Boolean
        Return TrySetPowerPointBestNarrativePlaceholderText(
            sld,
            text,
            New Integer() {2, 7, 17, 6},
            occurrence
        )
    End Function

    Private Shared Function RenderAutoPilotPowerPointTemplateTextSlide(sld As Object,
                                                                        slideObj As JObject,
                                                                        semanticLayout As String) As Boolean
        If sld Is Nothing OrElse slideObj Is Nothing Then Return False

        Dim title As String = If(slideObj.Value(Of String)("title"), "")
        Dim subtitle As String = If(slideObj.Value(Of String)("subtitle"), "")
        Dim body As String = If(slideObj.Value(Of String)("body"), "")
        Dim titleApplied As Boolean = TrySetPowerPointTemplateTitle(sld, title)

        ' Only populate a subtitle when the customer template exposes a genuine subtitle
        ' placeholder. Never consume a body/object placeholder merely to display a subtitle.
        If subtitle <> "" AndAlso semanticLayout <> "title" AndAlso semanticLayout <> "section" AndAlso semanticLayout <> "closing" Then
            TrySetPowerPointPlaceholderText(sld, subtitle, New Integer() {4}, 0)
        End If

        Select Case semanticLayout
            Case "title"
                Dim secondLine As String = If(subtitle <> "", subtitle, body)
                Dim contentApplied As Boolean = TrySetPowerPointPlaceholderText(
                    sld,
                    secondLine,
                    New Integer() {4, 2, 7, 17, 6},
                    0
                )
                Return titleApplied AndAlso (String.IsNullOrWhiteSpace(secondLine) OrElse contentApplied)

            Case "section"
                Dim sectionNumber As String = If(slideObj.Value(Of String)("section_number"), "")
                If sectionNumber <> "" Then
                    TrySetPowerPointPlaceholderText(sld, sectionNumber, New Integer() {2, 6, 7, 17}, 0)
                End If
                If subtitle <> "" Then
                    If Not TrySetPowerPointPlaceholderText(sld, subtitle, New Integer() {4}, 0) Then
                        TrySetPowerPointPlaceholderText(sld, subtitle, New Integer() {2, 6, 7, 17}, If(sectionNumber <> "", 1, 0))
                    End If
                End If
                Return titleApplied

            Case "bullets"
                Dim bulletItems As JArray = GetAutoPilotPowerPointBulletItems(slideObj)
                Dim contentText As String = CleanPptBulletText(body)
                Dim contentApplied As Boolean
                Dim hasStructuredBullets As Boolean = bulletItems IsNot Nothing AndAlso bulletItems.Count > 0
                If hasStructuredBullets Then
                    contentApplied = TrySetPowerPointBestNarrativePlaceholderBulletItems(sld, bulletItems, New Integer() {2, 7, 17, 6}, 0)
                Else
                    contentApplied = TrySetPowerPointTemplateContent(sld, contentText, 0)
                End If
                Dim contentRequired As Boolean = hasStructuredBullets OrElse Not String.IsNullOrWhiteSpace(contentText)
                Return titleApplied AndAlso ((Not contentRequired) OrElse contentApplied)

            Case "two_column"
                Dim leftText As String = JoinPowerPointTextParts(
                    slideObj.Value(Of String)("left_title"),
                    slideObj.Value(Of String)("left_body")
                )
                Dim rightText As String = JoinPowerPointTextParts(
                    slideObj.Value(Of String)("right_title"),
                    slideObj.Value(Of String)("right_body")
                )
                Dim leftApplied As Boolean = TrySetPowerPointTemplateContent(sld, leftText, 0)
                Dim rightApplied As Boolean = TrySetPowerPointTemplateContent(sld, rightText, 1)
                Return titleApplied AndAlso leftApplied AndAlso rightApplied

            Case "quote"
                Dim quoteText As String = If(slideObj.Value(Of String)("quote"), "")
                If quoteText = "" Then quoteText = body
                Dim attribution As String = If(slideObj.Value(Of String)("attribution"), "")
                Dim combined As String = JoinPowerPointTextParts(quoteText, attribution)
                Dim contentApplied As Boolean = TrySetPowerPointTemplateContent(sld, combined, 0)
                Return (titleApplied OrElse String.IsNullOrWhiteSpace(title)) AndAlso contentApplied

            Case "closing"
                Dim closingText As String = If(subtitle <> "", subtitle, body)
                Dim contentApplied As Boolean = True
                If closingText <> "" Then
                    contentApplied = TrySetPowerPointPlaceholderText(
                        sld,
                        closingText,
                        New Integer() {4, 2, 6, 7, 17},
                        0
                    )
                End If
                Return titleApplied AndAlso contentApplied

            Case Else
                Return False
        End Select
    End Function

    Private Shared Function JoinPowerPointTextParts(first As String, second As String) As String
        Dim parts As New List(Of String)()
        If Not String.IsNullOrWhiteSpace(first) Then parts.Add(first.Trim())
        If Not String.IsNullOrWhiteSpace(second) Then parts.Add(CleanPptBulletText(second))
        Return String.Join(System.Environment.NewLine, parts)
    End Function

    Private Shared Function BuildPowerPointFallbackBody(slideObj As JObject) As String
        If slideObj Is Nothing Then Return ""
        Dim body As String = CleanPptBulletText(slideObj.Value(Of String)("body"))
        If Not String.IsNullOrWhiteSpace(body) Then Return body

        Dim lines As New List(Of String)()
        Dim layout As String = If(slideObj.Value(Of String)("layout"), "").Trim().ToLowerInvariant()

        If layout = "two_column" Then
            Dim leftText As String = JoinPowerPointTextParts(slideObj.Value(Of String)("left_title"), slideObj.Value(Of String)("left_body"))
            Dim rightText As String = JoinPowerPointTextParts(slideObj.Value(Of String)("right_title"), slideObj.Value(Of String)("right_body"))
            If leftText <> "" Then lines.Add(leftText)
            If rightText <> "" Then lines.Add(rightText)
        End If

        For Each arrayName As String In New String() {"kpis", "cards", "steps", "events"}
            Dim items As JArray = TryCast(slideObj(arrayName), JArray)
            If items Is Nothing Then Continue For
            Dim itemNumber As Integer = 0
            For Each item As JObject In items.OfType(Of JObject)()
                itemNumber += 1
                Dim lead As String = If(item.Value(Of String)("label"), "")
                If lead = "" Then lead = If(item.Value(Of String)("title"), "")
                If lead = "" Then lead = If(item.Value(Of String)("value"), "")
                Dim detail As String = If(item.Value(Of String)("body"), "")
                If detail = "" Then detail = If(item.Value(Of String)("subtitle"), "")
                Dim line As String = JoinPowerPointTextParts(lead, detail)
                If line <> "" Then
                    If arrayName = "steps" AndAlso lead = "" Then line = $"{itemNumber}. {line}"
                    lines.Add(line)
                End If
            Next
        Next

        Dim structureObj As JObject = TryCast(slideObj("structure"), JObject)
        If structureObj IsNot Nothing Then
            Dim hierarchyNodes As JArray = TryCast(structureObj("nodes"), JArray)
            If hierarchyNodes IsNot Nothing Then
                For Each nodeObj As JObject In hierarchyNodes.OfType(Of JObject)()
                    Dim nodeText As String = JoinPowerPointTextParts(nodeObj.Value(Of String)("label"), nodeObj.Value(Of String)("detail"))
                    If nodeText <> "" Then lines.Add(nodeText)
                Next
            End If
            Dim topNode As JObject = TryCast(structureObj("top"), JObject)
            If topNode IsNot Nothing Then
                Dim topText As String = JoinPowerPointTextParts(topNode.Value(Of String)("title"), topNode.Value(Of String)("body"))
                If topText <> "" Then lines.Add(topText)
            End If
            Dim children As JArray = TryCast(structureObj("children"), JArray)
            If children IsNot Nothing Then
                For Each child As JObject In children.OfType(Of JObject)()
                    Dim childText As String = JoinPowerPointTextParts(child.Value(Of String)("title"), child.Value(Of String)("body"))
                    If childText <> "" Then lines.Add(childText)
                Next
            End If
        End If

        Dim comparison As JObject = TryCast(slideObj("comparison"), JObject)
        If comparison IsNot Nothing Then
            Dim columns As JArray = TryCast(comparison("columns"), JArray)
            If columns IsNot Nothing Then
                For Each column As JObject In columns.OfType(Of JObject)()
                    Dim columnLines As New List(Of String)()
                    Dim columnTitle As String = If(column.Value(Of String)("title"), "")
                    If columnTitle <> "" Then columnLines.Add(columnTitle)
                    Dim items As JArray = TryCast(column("items"), JArray)
                    If items IsNot Nothing Then
                        For Each token As JToken In items
                            If token IsNot Nothing AndAlso token.Type = JTokenType.String Then columnLines.Add(token.ToString())
                        Next
                    End If
                    Dim verdict As String = If(column.Value(Of String)("verdict"), "")
                    If verdict <> "" Then columnLines.Add(verdict)
                    If columnLines.Count > 0 Then lines.Add(String.Join(" — ", columnLines))
                Next
            End If
        End If

        Dim matrix As JObject = TryCast(slideObj("matrix"), JObject)
        If matrix IsNot Nothing Then
            Dim quadrants As JArray = TryCast(matrix("quadrants"), JArray)
            If quadrants IsNot Nothing Then
                For Each quadrant As JObject In quadrants.OfType(Of JObject)()
                    Dim quadrantText As String = JoinPowerPointTextParts(quadrant.Value(Of String)("title"), quadrant.Value(Of String)("body"))
                    If quadrantText <> "" Then lines.Add(quadrantText)
                Next
            End If
        End If

        Return String.Join(System.Environment.NewLine, lines)
    End Function

    Private Shared Sub RenderPowerPointFallbackContent(sld As Object,
                                                        slideObj As JObject,
                                                        templateLayoutApplied As Boolean,
                                                        slideW As Single,
                                                        fontName As String,
                                                        textColor As Integer,
                                                        light As Integer,
                                                        lineColor As Integer,
                                                        accent As Integer)
        Dim fallbackBody As String = BuildPowerPointFallbackBody(slideObj)
        If String.IsNullOrWhiteSpace(fallbackBody) Then Return

        If templateLayoutApplied AndAlso TrySetPowerPointTemplateContent(sld, fallbackBody, 0) Then Return
        RenderPptBullets(sld, fallbackBody, slideW, fontName, textColor, light, lineColor, accent)
    End Sub

    Private Shared Sub ApplyPowerPointTemplateFooter(sld As Object,
                                                      footerText As String,
                                                      slideIndex As Integer,
                                                      showSlideNumbers As Boolean)
        If sld Is Nothing Then Exit Sub
        If Not String.IsNullOrWhiteSpace(footerText) Then
            TrySetPowerPointPlaceholderText(sld, footerText, New Integer() {15}, 0)
        End If

        If showSlideNumbers Then
            Try
                sld.HeadersFooters.SlideNumber.Visible = -1
            Catch ex As System.Exception
                TrySetPowerPointPlaceholderText(sld, slideIndex.ToString(), New Integer() {13}, 0)
            End Try
        End If
    End Sub


    Private Shared Function GetAutoPilotPowerPointDataPayload(slideObj As JObject) As JObject
        If slideObj Is Nothing Then Return Nothing
        Dim dataToken As JToken = slideObj("data_json")
        If dataToken Is Nothing OrElse dataToken.Type = JTokenType.Null Then Return Nothing

        Try
            If dataToken.Type = JTokenType.Object Then
                Return DirectCast(dataToken, JObject)
            End If

            Dim rawData As String = dataToken.ToString()
            If String.IsNullOrWhiteSpace(rawData) Then Return Nothing
            Return JObject.Parse(rawData)
        Catch ex As Newtonsoft.Json.JsonException
            Debug.WriteLine($"PowerPoint data_json parse error: {ex.Message}")
            Return Nothing
        Catch ex As System.Exception
            Debug.WriteLine($"PowerPoint data_json expansion error: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Shared Function HasAutoPilotPowerPointStructuredPayloadInDataJson(slideObj As JObject, semanticLayout As String) As Boolean
        Dim payload As JObject = GetAutoPilotPowerPointDataPayload(slideObj)
        If payload Is Nothing Then Return False

        Select Case If(semanticLayout, "").Trim().ToLowerInvariant()
            Case "kpi"
                Dim items As JArray = TryCast(payload("kpis"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "table"
                Return TryCast(payload("table"), JObject) IsNot Nothing
            Case "chart"
                Return TryCast(payload("chart"), JObject) IsNot Nothing
            Case "cards"
                Dim items As JArray = TryCast(payload("cards"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "process"
                Dim items As JArray = TryCast(payload("steps"), JArray)
                Return items IsNot Nothing AndAlso items.Count > 0
            Case "structure"
                Return TryCast(payload("structure"), JObject) IsNot Nothing
            Case "timeline"
                Dim events As JArray = TryCast(payload("events"), JArray)
                If events Is Nothing Then events = TryCast(payload("timeline"), JArray)
                If events Is Nothing Then
                    Dim timelineObject As JObject = TryCast(payload("timeline"), JObject)
                    If timelineObject IsNot Nothing Then events = TryCast(timelineObject("events"), JArray)
                End If
                Return events IsNot Nothing AndAlso events.Count > 0
            Case "comparison"
                Return TryCast(payload("comparison"), JObject) IsNot Nothing
            Case "matrix"
                Return TryCast(payload("matrix"), JObject) IsNot Nothing
        End Select

        Return False
    End Function

    Private Shared Sub ExpandAutoPilotPowerPointSlideData(slideObj As JObject)
        If slideObj Is Nothing Then Exit Sub

        Try
            ' Preferred typed contracts are expanded first. This mirrors the successful Word visual
            ' contract: the model describes semantic structure and the host owns deterministic rendering.
            ExpandAutoPilotPowerPointTypedVisual(slideObj)

            Dim bulletItems As JArray = GetAutoPilotPowerPointBulletItems(slideObj)
            If bulletItems IsNot Nothing AndAlso bulletItems.Count > 0 Then
                slideObj("bullet_items") = bulletItems
                If String.IsNullOrWhiteSpace(slideObj.Value(Of String)("body")) Then
                    slideObj("body") = BuildAutoPilotPowerPointBulletBody(bulletItems)
                End If
            End If

            Dim payload As JObject = GetAutoPilotPowerPointDataPayload(slideObj)
            If payload Is Nothing Then Exit Sub

            ' Be tolerant of a provider wrapping a structured payload once more with its semantic name,
            ' e.g. {"process":{"steps":[...]}} instead of the canonical {"steps":[...]}.
            Dim wrappedProcess As JObject = TryCast(payload("process"), JObject)
            If wrappedProcess IsNot Nothing AndAlso payload("steps") Is Nothing AndAlso wrappedProcess("steps") IsNot Nothing Then
                payload("steps") = wrappedProcess("steps").DeepClone()
            End If
            Dim wrappedKpi As JObject = TryCast(payload("kpi"), JObject)
            If wrappedKpi IsNot Nothing AndAlso payload("kpis") Is Nothing AndAlso wrappedKpi("kpis") IsNot Nothing Then
                payload("kpis") = wrappedKpi("kpis").DeepClone()
            End If
            Dim wrappedCards As JObject = TryCast(payload("cards"), JObject)
            If wrappedCards IsNot Nothing AndAlso wrappedCards("cards") IsNot Nothing Then
                payload("cards") = wrappedCards("cards").DeepClone()
            End If

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

    Private Structure AutoPilotPowerPointInteropRect
        Public X As Single
        Public Y As Single
        Public W As Single
        Public H As Single
    End Structure

    Private Shared Function GetAutoPilotPowerPointInteropRichCanvas(sld As Object,
                                                                     slideW As Single,
                                                                     slideH As Single,
                                                                     settings As Dictionary(Of String, String),
                                                                     ByRef zoneCount As Integer) As AutoPilotPowerPointInteropRect
        Dim zones As New List(Of AutoPilotPowerPointInteropRect)()
        zoneCount = 0
        Try
            Dim shapes As Object = sld.Shapes
            Try
                Dim count As Integer = CInt(shapes.Count)
                For i As Integer = 1 To count
                    Dim shape As Object = Nothing
                    Try
                        shape = shapes(i)
                        Dim phType As Integer = 0
                        Try : phType = CInt(shape.PlaceholderFormat.Type) : Catch : Continue For : End Try
                        If phType = 1 OrElse phType = 3 OrElse phType = 4 OrElse phType = 5 OrElse phType = 13 OrElse phType = 14 OrElse phType = 15 OrElse phType = 16 Then Continue For
                        If Not (phType = 2 OrElse phType = 6 OrElse phType = 7 OrElse phType = 8 OrElse phType = 9 OrElse phType = 11 OrElse phType = 12 OrElse phType = 17 OrElse phType = 18) Then Continue For
                        Dim rect As New AutoPilotPowerPointInteropRect With {
                            .X = CSng(shape.Left), .Y = CSng(shape.Top), .W = CSng(shape.Width), .H = CSng(shape.Height)
                        }
                        If rect.W < slideW * 0.18F OrElse rect.H < slideH * 0.20F Then Continue For
                        If rect.Y < slideH * 0.12F AndAlso rect.H < slideH * 0.45F Then Continue For
                        zones.Add(rect)
                    Finally
                        If shape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch : End Try
                    End Try
                Next
            Finally
                If shapes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch : End Try
            End Try
        Catch ex As System.Exception
        End Try

        Dim result As AutoPilotPowerPointInteropRect
        If zones.Count > 0 Then
            Dim left As Single = zones.Min(Function(r) r.X)
            Dim top As Single = zones.Min(Function(r) r.Y)
            Dim right As Single = zones.Max(Function(r) r.X + r.W)
            Dim bottom As Single = zones.Max(Function(r) r.Y + r.H)
            result = New AutoPilotPowerPointInteropRect With {.X = left, .Y = top, .W = right - left, .H = bottom - top}
            zoneCount = zones.Count
        Else
            result = New AutoPilotPowerPointInteropRect With {
                .X = slideW * 0.06F, .Y = slideH * 0.23F, .W = slideW * 0.88F, .H = slideH * 0.64F
            }
        End If

        Dim insetPct As Single = CSng(Math.Max(0.0R, Math.Min(10.0R, GetPowerPointGuidanceSettingDouble(settings, "rich.canvas_inset_pct", 2.0R))) / 100.0R)
        Dim insetX As Single = result.W * insetPct
        Dim insetY As Single = result.H * insetPct
        result.X += insetX : result.Y += insetY
        result.W = Math.Max(10.0F, result.W - 2.0F * insetX)
        result.H = Math.Max(10.0F, result.H - 2.0F * insetY)
        Return result
    End Function

    Private Shared Function TryFitAutoPilotPowerPointRichShapes(sld As Object,
                                                                 firstNewShapeIndex As Integer,
                                                                 canvas As AutoPilotPowerPointInteropRect,
                                                                 minimumDetailPt As Single,
                                                                 minimumLabelPt As Single,
                                                                 context As ToolExecutionContext,
                                                                 slideIndex As Integer,
                                                                 semanticLayout As String) As Boolean
        Dim shapes As Object = Nothing
        Try
            shapes = sld.Shapes
            Dim count As Integer = CInt(shapes.Count)
            If firstNewShapeIndex > count Then Return False

            Dim srcLeft As Single = Single.MaxValue, srcTop As Single = Single.MaxValue
            Dim srcRight As Single = Single.MinValue, srcBottom As Single = Single.MinValue
            For i As Integer = firstNewShapeIndex To count
                Dim shape As Object = Nothing
                Try
                    shape = shapes(i)
                    srcLeft = Math.Min(srcLeft, CSng(shape.Left))
                    srcTop = Math.Min(srcTop, CSng(shape.Top))
                    srcRight = Math.Max(srcRight, CSng(shape.Left + shape.Width))
                    srcBottom = Math.Max(srcBottom, CSng(shape.Top + shape.Height))
                Finally
                    If shape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch : End Try
                End Try
            Next
            Dim srcW As Single = Math.Max(1.0F, srcRight - srcLeft)
            Dim srcH As Single = Math.Max(1.0F, srcBottom - srcTop)
            Dim scale As Single = Math.Min(1.0F, Math.Min(canvas.W / srcW, canvas.H / srcH))
            If scale <= 0.0F Then Return False

            ' Small geometric reductions are disproportionately harmful to editable PowerPoint
            ' text boxes: PowerPoint can reflow a previously one-line label into two lines while
            ' the box height is reduced at the same time. For near-fit compositions preserve the
            ' renderer's native shape dimensions and only recenter them. The visual may use a few
            ' points beyond the abstract content canvas, but remains inside the slide's protected
            ' title/footer margins. Larger mismatches are still scaled normally.
            Dim transformScale As Single = If(scale >= 0.9F, 1.0F, scale)
            Dim targetLeft As Single = canvas.X + (canvas.W - srcW * transformScale) / 2.0F
            Dim targetTop As Single = canvas.Y + (canvas.H - srcH * transformScale) / 2.0F
            If context IsNot Nothing AndAlso Math.Abs(transformScale - scale) > 0.001F Then
                context.Log($"PowerPoint rich Interop near-fit preserved native shape size: slide={slideIndex}; semantic='{semanticLayout}'; computedScale={scale:F3}; appliedScale={transformScale:F3}.", "diag")
            End If

            For i As Integer = firstNewShapeIndex To count
                Dim shape As Object = Nothing
                Try
                    shape = shapes(i)
                    Dim oldLeft As Single = CSng(shape.Left)
                    Dim oldTop As Single = CSng(shape.Top)
                    Dim oldW As Single = CSng(shape.Width)
                    Dim oldH As Single = CSng(shape.Height)
                    shape.Left = targetLeft + (oldLeft - srcLeft) * transformScale
                    shape.Top = targetTop + (oldTop - srcTop) * transformScale
                    shape.Width = Math.Max(1.0F, oldW * transformScale)
                    shape.Height = Math.Max(1.0F, oldH * transformScale)
                Finally
                    If shape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch : End Try
                End Try
            Next

            ' Let PowerPoint measure the final rendered text. Use the real TextFrame2 margins instead
            ' of subtracting arbitrary padding. Only narrative/title boxes participate in the configured
            ' minimum-font rule; compact badges, step numbers and other labels keep their intentional size.
            For i As Integer = firstNewShapeIndex To count
                Dim shape As Object = Nothing
                Try
                    shape = shapes(i)
                    Dim hasText As Boolean = False
                    Try : hasText = CInt(shape.HasTextFrame) <> 0 AndAlso CInt(shape.TextFrame2.HasText) <> 0 : Catch : End Try
                    If Not hasText Then Continue For
                    Dim textRange As Object = Nothing
                    Dim textFrame2 As Object = Nothing
                    Try
                        textFrame2 = shape.TextFrame2
                        textRange = textFrame2.TextRange
                        Dim currentPt As Single = CSng(textRange.Font.Size)
                        Dim shapeHeight As Single = CSng(shape.Height)
                        Dim isBold As Boolean = False
                        Try : isBold = CSng(textRange.Font.Bold) <> 0.0F : Catch : End Try

                        ' Rich-diagram labels/details are not slide titles or narrative body paragraphs.
                        ' Preserve the renderer's role-specific font sizes instead of inflating every bold
                        ' label to rich.min_title_font_pt or every detail to rich.min_body_font_pt.
                        Dim roleMinimum As Single = If(isBold, minimumLabelPt, minimumDetailPt)
                        If currentPt > 0.0F AndAlso currentPt >= 10.0F AndAlso shapeHeight >= 28.0F AndAlso
                           roleMinimum > 0.0F AndAlso currentPt < roleMinimum Then
                            textRange.Font.Size = roleMinimum
                        End If

                        Dim boundH As Single = CSng(textRange.BoundHeight)
                        Dim boundW As Single = CSng(textRange.BoundWidth)
                        Dim marginLeft As Single = 0.0F, marginRight As Single = 0.0F, marginTop As Single = 0.0F, marginBottom As Single = 0.0F
                        Try : marginLeft = CSng(textFrame2.MarginLeft) : Catch : End Try
                        Try : marginRight = CSng(textFrame2.MarginRight) : Catch : End Try
                        Try : marginTop = CSng(textFrame2.MarginTop) : Catch : End Try
                        Try : marginBottom = CSng(textFrame2.MarginBottom) : Catch : End Try
                        Dim availableH As Single = Math.Max(1.0F, CSng(shape.Height) - marginTop - marginBottom)
                        Dim availableW As Single = Math.Max(1.0F, CSng(shape.Width) - marginLeft - marginRight)

                        ' PowerPoint occasionally reports a text box with a much smaller effective
                        ' height after COM placement/reflow than the renderer requested. Before
                        ' discarding the entire editable visual, restore the individual text box to
                        ' the height/width actually required by TextFrame2 when that recovery remains
                        ' wholly inside the safe visual canvas. This changes geometry, not typography.
                        If boundH > availableH * 1.08F AndAlso boundW <= availableW * 1.08F Then
                            Dim requiredHeight As Single = boundH + marginTop + marginBottom + 2.0F
                            Dim maxHeight As Single = canvas.Y + canvas.H - CSng(shape.Top)
                            If requiredHeight > CSng(shape.Height) AndAlso requiredHeight <= maxHeight AndAlso requiredHeight <= CSng(shape.Height) + 96.0F Then
                                shape.Height = requiredHeight
                                availableH = Math.Max(1.0F, CSng(shape.Height) - marginTop - marginBottom)
                                If context IsNot Nothing Then context.Log($"PowerPoint rich Interop recovered text-box height: slide={slideIndex}; semantic='{semanticLayout}'; shape={i}; height={CSng(shape.Height):F1}.", "diag")
                            End If
                        End If
                        If boundW > availableW * 1.08F AndAlso boundH <= availableH * 1.08F Then
                            Dim requiredWidth As Single = boundW + marginLeft + marginRight + 2.0F
                            Dim maxWidth As Single = canvas.X + canvas.W - CSng(shape.Left)
                            If requiredWidth > CSng(shape.Width) AndAlso requiredWidth <= maxWidth AndAlso requiredWidth <= CSng(shape.Width) + 96.0F Then
                                shape.Width = requiredWidth
                                availableW = Math.Max(1.0F, CSng(shape.Width) - marginLeft - marginRight)
                                If context IsNot Nothing Then context.Log($"PowerPoint rich Interop recovered text-box width: slide={slideIndex}; semantic='{semanticLayout}'; shape={i}; width={CSng(shape.Width):F1}.", "diag")
                            End If
                        End If

                        Dim shapeText As String = ""
                        Try : shapeText = CStr(textRange.Text) : Catch : End Try
                        Dim compactText As String = If(shapeText, "").Trim()
                        Dim decorativeText As Boolean = compactText = "→" OrElse compactText = "←" OrElse compactText = "↔" OrElse compactText.Length <= 1
                        If (Not decorativeText) AndAlso (boundH > availableH * 1.08F OrElse boundW > availableW * 1.08F) Then
                            If shapeText.Length > 90 Then shapeText = shapeText.Substring(0, 90)
                            shapeText = shapeText.Replace(vbCr, " ").Replace(vbLf, " ")
                            If context IsNot Nothing Then context.Log($"PowerPoint rich Interop text-fit rejected: slide={slideIndex}; semantic='{semanticLayout}'; shape={i}; font={CSng(textRange.Font.Size):F1}; bound={boundW:F1}x{boundH:F1}; available={availableW:F1}x{availableH:F1}; text='{shapeText}'.", "diag")
                            Return False
                        End If
                    Finally
                        If textRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(textRange) : Catch : End Try
                        If textFrame2 IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(textFrame2) : Catch : End Try
                    End Try
                Finally
                    If shape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch : End Try
                End Try
            Next
            Return True
        Catch ex As System.Exception
            If context IsNot Nothing Then context.Log($"PowerPoint rich Interop fit exception: slide={slideIndex}; semantic='{semanticLayout}'; {ex.Message}", "diag")
            Return False
        Finally
            If shapes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch : End Try
        End Try
    End Function

    Private Shared Sub DeleteAutoPilotPowerPointShapesFromIndex(sld As Object, firstShapeIndex As Integer)
        Try
            For i As Integer = CInt(sld.Shapes.Count) To firstShapeIndex Step -1
                Try : sld.Shapes(i).Delete() : Catch : End Try
            Next
        Catch ex As System.Exception
        End Try
    End Sub

    Private Shared Function BuildAutoPilotPowerPointRichFallbackBulletItems(slideObj As JObject, semanticLayout As String) As JArray
        Dim result As New JArray()
        If slideObj Is Nothing Then Return result

        If String.Equals(semanticLayout, "comparison", StringComparison.OrdinalIgnoreCase) Then
            Dim comparisonObj As JObject = TryCast(slideObj("comparison"), JObject)
            Dim columns As JArray = If(comparisonObj Is Nothing, Nothing, TryCast(comparisonObj("columns"), JArray))
            If columns IsNot Nothing Then
                For Each columnObj As JObject In columns.OfType(Of JObject)()
                    Dim columnTitle As String = If(columnObj.Value(Of String)("title"), "").Trim()
                    If columnTitle <> "" Then result.Add(New JObject From {{"text", columnTitle}, {"level", 0}})
                    Dim items As JArray = TryCast(columnObj("items"), JArray)
                    If items IsNot Nothing Then
                        For Each itemToken As JToken In items
                            Dim itemText As String = If(itemToken Is Nothing, "", itemToken.ToString()).Trim()
                            If itemText <> "" Then result.Add(New JObject From {{"text", itemText}, {"level", 1}})
                        Next
                    End If
                    Dim verdict As String = If(columnObj.Value(Of String)("verdict"), "").Trim()
                    If verdict <> "" Then result.Add(New JObject From {{"text", verdict}, {"level", 1}})
                Next
            End If
        ElseIf String.Equals(semanticLayout, "process", StringComparison.OrdinalIgnoreCase) OrElse
               String.Equals(semanticLayout, "timeline", StringComparison.OrdinalIgnoreCase) Then
            Dim items As JArray = TryCast(slideObj("steps"), JArray)
            If items Is Nothing Then items = TryCast(slideObj("events"), JArray)
            If items Is Nothing Then items = TryCast(slideObj("timeline"), JArray)
            If items IsNot Nothing Then
                For Each itemObj As JObject In items.OfType(Of JObject)()
                    Dim label As String = If(itemObj.Value(Of String)("title"), "").Trim()
                    If label = "" Then label = If(itemObj.Value(Of String)("label"), "").Trim()
                    Dim detail As String = If(itemObj.Value(Of String)("body"), "").Trim()
                    If detail = "" Then detail = If(itemObj.Value(Of String)("detail"), "").Trim()
                    If label <> "" Then result.Add(New JObject From {{"text", label}, {"level", 0}})
                    If detail <> "" Then result.Add(New JObject From {{"text", detail}, {"level", 1}})
                Next
            End If
        End If
        Return result
    End Function

    Private Shared Function RenderAutoPilotPowerPointRichInterop(sld As Object,
                                                                  slideObj As JObject,
                                                                  slideW As Single,
                                                                  slideH As Single,
                                                                  semanticLayout As String,
                                                                  settings As Dictionary(Of String, String),
                                                                  context As ToolExecutionContext,
                                                                  slideIndex As Integer,
                                                                  renderAction As System.Action) As Boolean
        Dim firstNewShapeIndex As Integer = CInt(sld.Shapes.Count) + 1
        Dim zoneCount As Integer = 0
        Dim canvas As AutoPilotPowerPointInteropRect = GetAutoPilotPowerPointInteropRichCanvas(sld, slideW, slideH, settings, zoneCount)
        If context IsNot Nothing Then context.Log($"PowerPoint rich Interop stage: slide={slideIndex}; semantic='{semanticLayout}'; stage=render-start; canvas={canvas.X:F1},{canvas.Y:F1},{canvas.W:F1},{canvas.H:F1}; zones={zoneCount}.", "diag")
        Try
            renderAction()
            Dim configuredBodyPt As Single = CSng(GetPowerPointGuidanceSettingDouble(settings, "rich.min_body_font_pt", 14.0R))
            Dim minDetailPt As Single = CSng(GetPowerPointGuidanceSettingDouble(settings, "rich.min_detail_font_pt", Math.Max(11.0R, configuredBodyPt * 0.72R)))
            Dim minLabelPt As Single = CSng(GetPowerPointGuidanceSettingDouble(settings, "rich.min_label_font_pt", Math.Max(13.0R, minDetailPt + 2.0R)))
            If TryFitAutoPilotPowerPointRichShapes(sld, firstNewShapeIndex, canvas, minDetailPt, minLabelPt, context, slideIndex, semanticLayout) Then
                If context IsNot Nothing Then context.Log($"PowerPoint rich Interop stage: slide={slideIndex}; semantic='{semanticLayout}'; stage=complete.", "diag")
                Return True
            End If
        Catch ex As System.Exception
            If context IsNot Nothing Then context.Log($"PowerPoint rich Interop stage: slide={slideIndex}; semantic='{semanticLayout}'; stage=render-exception; error={ex.Message}", "diag")
        End Try

        DeleteAutoPilotPowerPointShapesFromIndex(sld, firstNewShapeIndex)
        Dim fallbackBody As String = BuildPowerPointFallbackBody(slideObj)
        Dim applied As Boolean = False

        ' Preserve the native geometry when a rich comparison falls back. A 2-Content layout must
        ' receive one option per placeholder; putting the entire comparison into occurrence 0 causes
        ' overflow and leaves the second column unused.
        If String.Equals(semanticLayout, "comparison", StringComparison.OrdinalIgnoreCase) Then
            Dim comparisonObj As JObject = TryCast(slideObj("comparison"), JObject)
            Dim columns As JArray = If(comparisonObj Is Nothing, Nothing, TryCast(comparisonObj("columns"), JArray))
            If columns IsNot Nothing AndAlso columns.Count >= 2 Then
                Dim leftObj As JObject = TryCast(columns(0), JObject)
                Dim rightObj As JObject = TryCast(columns(1), JObject)
                If leftObj IsNot Nothing AndAlso rightObj IsNot Nothing Then
                    Dim leftText As String = JoinPowerPointTextParts(leftObj.Value(Of String)("title"), GetPptArrayText(leftObj, "items"))
                    Dim rightText As String = JoinPowerPointTextParts(rightObj.Value(Of String)("title"), GetPptArrayText(rightObj, "items"))
                    applied = TrySetPowerPointTemplateContent(sld, leftText, 0) AndAlso TrySetPowerPointTemplateContent(sld, rightText, 1)
                    If applied Then
                        slideObj("layout") = "two_column"
                        slideObj("left_title") = leftObj.Value(Of String)("title")
                        slideObj("left_body") = GetPptArrayText(leftObj, "items")
                        slideObj("right_title") = rightObj.Value(Of String)("title")
                        slideObj("right_body") = GetPptArrayText(rightObj, "items")
                    End If
                End If
            End If
        End If

        If Not applied Then
            Dim fallbackBulletItems As JArray = BuildAutoPilotPowerPointRichFallbackBulletItems(slideObj, semanticLayout)
            If fallbackBulletItems IsNot Nothing AndAlso fallbackBulletItems.Count > 0 Then
                applied = TrySetPowerPointBestNarrativePlaceholderBulletItems(sld, fallbackBulletItems, New Integer() {2, 7, 17, 6}, 0)
                If applied Then
                    slideObj("layout") = "bullets"
                    slideObj("bullet_items") = fallbackBulletItems
                    slideObj("body") = BuildAutoPilotPowerPointBulletBody(fallbackBulletItems)
                End If
            End If
        End If

        If Not applied Then
            applied = TrySetPowerPointTemplateContent(sld, fallbackBody, 0)
            If applied Then
                ' The renderer has deliberately changed the delivered representation. Keep the validation
                ' contract in sync so a decorative rich-only label (for example a card badge) is not later
                ' required from the native text fallback.
                slideObj("layout") = "bullets"
                slideObj("body") = fallbackBody
            End If
        End If
        If context IsNot Nothing Then context.Log($"PowerPoint rich Interop fallback: slide={slideIndex}; semantic='{semanticLayout}'; nativeTextApplied={applied.ToString().ToLowerInvariant()}.", "diag")
        Return applied
    End Function

    Private Shared Sub DiagnoseAutoPilotPowerPointOpenXmlPackage(outputPath As String,
                                                                  context As ToolExecutionContext)
        Try
            Using document As DocumentFormat.OpenXml.Packaging.PresentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(outputPath, False)
                Dim validator As New DocumentFormat.OpenXml.Validation.OpenXmlValidator()
                Dim validationIssues = validator.Validate(document).Take(12).ToList()
                If validationIssues.Count = 0 Then
                    If context IsNot Nothing Then context.Log("PowerPoint OpenXML diagnostic validation found no schema issues.", "diag")
                    Exit Sub
                End If
                For Each validationIssue In validationIssues
                    Dim partUri As String = ""
                    Try : partUri = If(validationIssue.Part Is Nothing, "", validationIssue.Part.Uri.ToString()) : Catch : End Try
                    Dim validationPath As String = ""
                    Try : validationPath = If(validationIssue.Path Is Nothing, "", validationIssue.Path.XPath) : Catch : End Try
                    If context IsNot Nothing Then context.Log($"PowerPoint OpenXML diagnostic issue: part='{partUri}'; path='{validationPath}'; description='{validationIssue.Description}'.", "diag")
                Next
            End Using
        Catch ex As System.Exception
            If context IsNot Nothing Then context.Log($"PowerPoint OpenXML diagnostic validation skipped: {ex.Message}", "diag")
        End Try
    End Sub

    Private Shared Function CollectAutoPilotPowerPointInteropSlideText(sld As Object) As String
        Dim values As New List(Of String)()
        If sld Is Nothing Then Return ""
        Dim shapes As Object = Nothing
        Try
            shapes = sld.Shapes
            For i As Integer = 1 To CInt(shapes.Count)
                Dim shape As Object = Nothing
                Try
                    shape = shapes(i)
                    Dim hasText As Boolean = False
                    Try : hasText = CInt(shape.HasTextFrame) <> 0 AndAlso CInt(shape.TextFrame.HasText) <> 0 : Catch : End Try
                    If hasText Then
                        Dim value As String = ""
                        Try : value = CStr(shape.TextFrame.TextRange.Text) : Catch : End Try
                        If Not String.IsNullOrWhiteSpace(value) Then values.Add(value)
                    End If
                Finally
                    If shape IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shape) : Catch : End Try
                End Try
            Next
        Finally
            If shapes IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(shapes) : Catch : End Try
        End Try
        Return String.Join(" ", values)
    End Function

    Private Shared Function ValidateAutoPilotPowerPointLivePresentation(pres As Object,
                                                                          slidesArray As JArray,
                                                                          templateAssignments As IDictionary(Of Integer, AutoPilotPowerPointLayoutCandidate),
                                                                          context As ToolExecutionContext) As String
        If pres Is Nothing OrElse slidesArray Is Nothing OrElse slidesArray.Count = 0 Then
            Return "PowerPoint output validation could not inspect the generated presentation."
        End If

        Try
            Dim totalSlides As Integer = CInt(pres.Slides.Count)
            If totalSlides < slidesArray.Count Then
                Return $"PowerPoint output validation expected {slidesArray.Count} generated slides but found only {totalSlides} total slides in the saved presentation."
            End If

            Dim firstGeneratedIndex As Integer = totalSlides - slidesArray.Count + 1
            For ordinal As Integer = 1 To slidesArray.Count
                Dim slideObj As JObject = TryCast(slidesArray(ordinal - 1), JObject)
                If slideObj Is Nothing Then Continue For
                Dim sld As Object = Nothing
                Dim customLayout As Object = Nothing
                Try
                    sld = pres.Slides(firstGeneratedIndex + ordinal - 1)
                    If templateAssignments IsNot Nothing AndAlso templateAssignments.ContainsKey(ordinal) Then
                        Dim expectedLayout As AutoPilotPowerPointLayoutCandidate = templateAssignments(ordinal)
                        Dim actualLayoutName As String = ""
                        Try
                            customLayout = sld.CustomLayout
                            actualLayoutName = CStr(customLayout.Name)
                        Catch
                        End Try
                        If Not String.IsNullOrWhiteSpace(expectedLayout.LayoutName) AndAlso
                           Not System.Text.RegularExpressions.Regex.IsMatch(expectedLayout.LayoutName.Trim(), "^Layout\s+\d+$", System.Text.RegularExpressions.RegexOptions.IgnoreCase) AndAlso
                           NormalizePowerPointLayoutKey(actualLayoutName) <> NormalizePowerPointLayoutKey(expectedLayout.LayoutName) Then
                            Return $"PowerPoint output validation failed on generated slide {ordinal}: expected template layout '{expectedLayout.LayoutName}' but the live saved presentation uses '{actualLayoutName}'."
                        End If
                    End If

                    Dim normalizedActual As String = NormalizeAutoPilotPowerPointValidationText(CollectAutoPilotPowerPointInteropSlideText(sld))
                    Dim fragments As List(Of String) = GetAutoPilotPowerPointValidationFragments(slideObj, ordinal)
                    For Each fragment As String In fragments
                        Dim expected As String = NormalizeAutoPilotPowerPointValidationText(fragment)
                        If expected = "" Then Continue For
                        If Not normalizedActual.Contains(expected) Then
                            Return $"PowerPoint output validation failed on generated slide {ordinal}: expected visible text was not found ('{fragment}'). The presentation was not registered for delivery."
                        End If
                    Next
                Finally
                    If customLayout IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(customLayout) : Catch : End Try
                    If sld IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(sld) : Catch : End Try
                End Try
            Next

            If context IsNot Nothing Then context.Log($"PowerPoint live validation completed successfully after SaveAs: slides={slidesArray.Count}.", "diag")
            Return ""
        Catch ex As System.Exception
            Dim validationError As String = $"PowerPoint output validation failed in the live saved presentation: {ex.Message}"
            If context IsNot Nothing Then context.Log(validationError, "diag")
            Return validationError
        End Try
    End Function

    Private Shared Sub RenderAutoPilotPowerPointSlide(pres As Object,
                                                        sld As Object,
                                                        slideObj As JObject,
                                                        slideIndex As Integer,
                                                        existingSlideCount As Integer,
                                                        args As Dictionary(Of String, Object),
                                                        templateLayoutApplied As Boolean,
                                                        templateLayout As AutoPilotPowerPointLayoutCandidate,
                                                        richSettings As Dictionary(Of String, String),
                                                        context As ToolExecutionContext)
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

        Dim layout As String = NormalizeAutoPilotPowerPointSemanticLayout(
            slideObj.Value(Of String)("layout"),
            slideIndex,
            existingSlideCount
        )

        Dim title As String = slideObj.Value(Of String)("title")
        Dim subtitle As String = slideObj.Value(Of String)("subtitle")
        Dim body As String = slideObj.Value(Of String)("body")
        Dim sourceText As String = slideObj.Value(Of String)("source")

        If templateLayoutApplied Then
            Try
                sld.FollowMasterBackground = -1
            Catch ex As System.Exception
            End Try

            Dim usesOpenXmlRichOverlay As Boolean =
                (layout = "cards" AndAlso TryCast(slideObj("cards"), JArray) IsNot Nothing) OrElse
                (layout = "comparison" AndAlso TryCast(slideObj("comparison"), JObject) IsNot Nothing) OrElse
                (layout = "structure" AndAlso TryCast(slideObj("structure"), JObject) IsNot Nothing) OrElse
                (layout = "process" AndAlso TryCast(slideObj("steps"), JArray) IsNot Nothing) OrElse
                (layout = "timeline" AndAlso (TryCast(slideObj("events"), JArray) IsNot Nothing OrElse TryCast(slideObj("timeline"), JArray) IsNot Nothing OrElse TryCast(slideObj("timeline"), JObject) IsNot Nothing)) OrElse
                (layout = "matrix" AndAlso TryCast(slideObj("matrix"), JObject) IsNot Nothing) OrElse
                (layout = "chart" AndAlso TryCast(slideObj("chart"), JObject) IsNot Nothing) OrElse
                (layout = "kpi" AndAlso TryCast(slideObj("kpis"), JArray) IsNot Nothing)

            ' Rich layouts intentionally continue into the Interop renderer below.
            ' The native layout/master remains applied, but PowerPoint itself creates the
            ' freeform graphical shapes. This avoids hand-written PresentationML.
            ' Deterministic semantic placement is the primary contract for template-backed
            ' text slides. Interpreter-generated bindings are only a fallback. This prevents
            ' a customer template's small Body-typed overtitle or metadata placeholder from
            ' receiving the narrative merely because its COM type happens to match first.
            If Not usesOpenXmlRichOverlay AndAlso RenderAutoPilotPowerPointTemplateTextSlide(sld, slideObj, layout) Then
                If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then
                    Dim templateFooterText As String = GetArgString(args, "footer_text")
                    If Not String.IsNullOrWhiteSpace(sourceText) Then templateFooterText = sourceText
                    ApplyPowerPointTemplateFooter(
                        sld,
                        templateFooterText,
                        slideIndex,
                        GetArgBool(args, "show_slide_numbers", True)
                    )
                End If
                Exit Sub
            End If

            If (Not usesOpenXmlRichOverlay) AndAlso templateLayout IsNot Nothing AndAlso
               templateLayout.TextBindings IsNot Nothing AndAlso
               templateLayout.TextBindings.Count > 0 Then

                If Not RenderAutoPilotPowerPointNativeTextBindings(sld, templateLayout) Then
                    Throw New System.Exception(
                        $"The selected PowerPoint template layout '{templateLayout.LayoutName}' could not receive its validated native placeholder bindings.")
                End If

                If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then
                    Dim templateFooterText As String = GetArgString(args, "footer_text")
                    If Not String.IsNullOrWhiteSpace(sourceText) Then templateFooterText = sourceText
                    ApplyPowerPointTemplateFooter(
                        sld,
                        templateFooterText,
                        slideIndex,
                        GetArgBool(args, "show_slide_numbers", True)
                    )
                End If
                Exit Sub
            End If

            If RenderAutoPilotPowerPointTemplateTextSlide(sld, slideObj, layout) Then
                If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then
                    Dim templateFooterText As String = GetArgString(args, "footer_text")
                    If Not String.IsNullOrWhiteSpace(sourceText) Then templateFooterText = sourceText
                    ApplyPowerPointTemplateFooter(
                        sld,
                        templateFooterText,
                        slideIndex,
                        GetArgBool(args, "show_slide_numbers", True)
                    )
                End If
                Exit Sub
            End If
        Else
            If layout = "title" OrElse layout = "section" OrElse layout = "closing" Then
                SetPptSlideBackground(sld, accent)
            Else
                SetPptSlideBackground(sld, PptHexColor("#FFFFFF", "#FFFFFF"))
                AddPptShape(sld, 1, 0.0F, 0.0F, slideW, 5.0F, secondary, secondary, 0.0F)
            End If
        End If

        Dim addStandardTitle As System.Action =
            Sub()
                If templateLayoutApplied Then
                    If Not TrySetPowerPointTemplateTitle(sld, title) Then
                        AddPptTextBox(sld, title, 52.0F, 36.0F, slideW - 104.0F, 58.0F, 24.0F, True, textColor, fontName, 1, 0.0F)
                    End If
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        ' On template-backed content/rich slides, only use a genuine subtitle placeholder.
                        ' Do not synthesize a free subtitle box: it can overlap the template canvas or appear
                        ' as stray text in layouts that intentionally have no subtitle region.
                        TrySetPowerPointPlaceholderText(sld, subtitle, New Integer() {4}, 0)
                    End If
                Else
                    AddPptStandardTitle(sld, title, subtitle, slideW, fontName, textColor, muted, accent)
                End If
            End Sub

        Select Case layout
            Case "title"
                If templateLayoutApplied Then
                    If Not TrySetPowerPointTemplateTitle(sld, title) Then
                        AddPptTextBox(sld, title, 58.0F, 142.0F, slideW - 116.0F, 132.0F, 35.0F, True, textColor, fontName, 1, 0.0F)
                    End If
                    Dim titleSecondLine As String = If(Not String.IsNullOrWhiteSpace(subtitle), subtitle, body)
                    If Not String.IsNullOrWhiteSpace(titleSecondLine) AndAlso
                       Not TrySetPowerPointPlaceholderText(sld, titleSecondLine, New Integer() {4, 2, 7, 17, 6}, 0) Then

                        AddPptTextBox(sld, titleSecondLine, 60.0F, 286.0F, slideW - 120.0F, 76.0F, 18.0F, False, textColor, fontName, 1, 0.0F)
                    End If
                Else
                    AddPptTextBox(sld, title, 58.0F, 142.0F, slideW - 116.0F, 132.0F, 35.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 1, 0.0F)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddPptTextBox(sld, subtitle, 60.0F, 286.0F, slideW - 120.0F, 76.0F, 18.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                    ElseIf Not String.IsNullOrWhiteSpace(body) Then
                        AddPptTextBox(sld, body, 60.0F, 286.0F, slideW - 120.0F, 76.0F, 18.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                    End If
                    AddPptShape(sld, 1, 60.0F, 118.0F, 72.0F, 6.0F, secondary, secondary, 0.0F)
                End If

            Case "section"
                If templateLayoutApplied Then
                    If Not TrySetPowerPointTemplateTitle(sld, title) Then
                        AddPptTextBox(sld, title, 60.0F, 166.0F, slideW - 120.0F, 124.0F, 32.0F, True, textColor, fontName, 1, 0.0F)
                    End If
                    Dim sectionNumber As String = slideObj.Value(Of String)("section_number")
                    If Not String.IsNullOrWhiteSpace(sectionNumber) Then
                        If Not TrySetPowerPointPlaceholderText(sld, sectionNumber, New Integer() {2, 6, 7, 17}, 0) Then
                            AddPptTextBox(sld, sectionNumber, 60.0F, 102.0F, 120.0F, 45.0F, 15.0F, True, accent, fontName, 1, 0.0F)
                        End If
                    End If
                    If Not String.IsNullOrWhiteSpace(subtitle) AndAlso
                       Not TrySetPowerPointPlaceholderText(sld, subtitle, New Integer() {4}, 0) Then

                        AddPptTextBox(sld, subtitle, 60.0F, 305.0F, slideW - 120.0F, 76.0F, 17.0F, False, muted, fontName, 1, 0.0F)
                    End If
                Else
                    Dim sectionNumber As String = slideObj.Value(Of String)("section_number")
                    If Not String.IsNullOrWhiteSpace(sectionNumber) Then
                        AddPptTextBox(sld, sectionNumber, 60.0F, 102.0F, 120.0F, 45.0F, 15.0F, True, PptHexColor("#BFD7EA", "#BFD7EA"), fontName, 1, 0.0F)
                    End If
                    AddPptTextBox(sld, title, 60.0F, 166.0F, slideW - 120.0F, 124.0F, 32.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 1, 0.0F)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then
                        AddPptTextBox(sld, subtitle, 60.0F, 305.0F, slideW - 120.0F, 76.0F, 17.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 1, 0.0F)
                    End If
                End If

            Case "two_column"
                addStandardTitle()
                Dim leftText As String = JoinPowerPointTextParts(slideObj.Value(Of String)("left_title"), slideObj.Value(Of String)("left_body"))
                Dim rightText As String = JoinPowerPointTextParts(slideObj.Value(Of String)("right_title"), slideObj.Value(Of String)("right_body"))
                If templateLayoutApplied AndAlso
                   TrySetPowerPointTemplateContent(sld, leftText, 0) AndAlso
                   TrySetPowerPointTemplateContent(sld, rightText, 1) Then
                Else
                    RenderPptTwoColumn(sld, slideObj, slideW, fontName, textColor, muted, light, lineColor, accent)
                End If

            Case "kpi"
                addStandardTitle()
                Dim kpis As JArray = TryCast(slideObj("kpis"), JArray)
                If kpis IsNot Nothing AndAlso kpis.Count > 0 Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptKpiCards(sld, kpis, slideW, fontName, textColor, muted, light, lineColor, accent))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "table"
                addStandardTitle()
                Dim tableObject As JObject = TryCast(slideObj("table"), JObject)
                If tableObject IsNot Nothing Then
                    RenderPptTable(sld, tableObject, slideW, fontName, textColor, muted, light, lineColor, accent)
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "chart"
                addStandardTitle()
                Dim chartObject As JObject = TryCast(slideObj("chart"), JObject)
                If chartObject IsNot Nothing Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptChart(sld, chartObject, slideW, fontName, textColor, muted, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If
                Dim callout As String = slideObj.Value(Of String)("callout")
                If Not String.IsNullOrWhiteSpace(callout) Then
                    AddPptTextBox(sld, callout, slideW - 330.0F, 448.0F, 280.0F, 48.0F, 12.0F, True, accent, fontName, 3, 0.0F)
                End If

            Case "cards"
                addStandardTitle()
                Dim cards As JArray = TryCast(slideObj("cards"), JArray)
                If cards IsNot Nothing AndAlso cards.Count > 0 Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptCards(sld, cards, slideW, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "process"
                addStandardTitle()
                Dim steps As JArray = TryCast(slideObj("steps"), JArray)
                If steps IsNot Nothing AndAlso steps.Count > 0 Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptProcess(sld, steps, slideW, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "structure"
                addStandardTitle()
                Dim structureObject As JObject = TryCast(slideObj("structure"), JObject)
                If structureObject IsNot Nothing Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptStructure(sld, structureObject, slideW, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "timeline"
                addStandardTitle()
                Dim events As JArray = TryCast(slideObj("events"), JArray)
                If events Is Nothing Then events = TryCast(slideObj("timeline"), JArray)
                If events Is Nothing Then
                    Dim timelineObject As JObject = TryCast(slideObj("timeline"), JObject)
                    If timelineObject IsNot Nothing Then events = TryCast(timelineObject("events"), JArray)
                End If
                If events IsNot Nothing AndAlso events.Count > 0 Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptTimeline(sld, events, slideW, slideH, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "comparison"
                addStandardTitle()
                Dim comparisonObject As JObject = TryCast(slideObj("comparison"), JObject)
                If comparisonObject IsNot Nothing Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptComparison(sld, comparisonObject, slideW, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "matrix"
                addStandardTitle()
                Dim matrixObject As JObject = TryCast(slideObj("matrix"), JObject)
                If matrixObject IsNot Nothing Then
                    RenderAutoPilotPowerPointRichInterop(sld, slideObj, slideW, slideH, layout, richSettings, context, slideIndex, Sub() RenderPptMatrix(sld, matrixObject, slideW, fontName, textColor, muted, light, lineColor, accent, secondary))
                Else
                    RenderPowerPointFallbackContent(sld, slideObj, templateLayoutApplied, slideW, fontName, textColor, light, lineColor, accent)
                End If

            Case "quote"
                addStandardTitle()
                Dim quoteText As String = slideObj.Value(Of String)("quote")
                If String.IsNullOrWhiteSpace(quoteText) Then quoteText = body
                AddPptTextBox(sld, "“" & quoteText & "”", 105.0F, 155.0F, slideW - 210.0F, 200.0F, 26.0F, False, accent, fontName, 2, 0.0F)
                Dim attribution As String = slideObj.Value(Of String)("attribution")
                If Not String.IsNullOrWhiteSpace(attribution) Then AddPptTextBox(sld, attribution, 160.0F, 375.0F, slideW - 320.0F, 45.0F, 13.0F, True, muted, fontName, 2, 0.0F)

            Case "closing"
                If templateLayoutApplied Then
                    If Not TrySetPowerPointTemplateTitle(sld, title) Then
                        AddPptTextBox(sld, title, 60.0F, 174.0F, slideW - 120.0F, 110.0F, 34.0F, True, textColor, fontName, 2, 0.0F)
                    End If
                    Dim closingText As String = If(Not String.IsNullOrWhiteSpace(subtitle), subtitle, body)
                    If Not String.IsNullOrWhiteSpace(closingText) AndAlso
                       Not TrySetPowerPointPlaceholderText(sld, closingText, New Integer() {4, 2, 6, 7, 17}, 0) Then

                        AddPptTextBox(sld, closingText, 80.0F, 300.0F, slideW - 160.0F, 76.0F, 17.0F, False, muted, fontName, 2, 0.0F)
                    End If
                Else
                    AddPptTextBox(sld, title, 60.0F, 174.0F, slideW - 120.0F, 110.0F, 34.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 2, 0.0F)
                    If Not String.IsNullOrWhiteSpace(subtitle) Then AddPptTextBox(sld, subtitle, 80.0F, 300.0F, slideW - 160.0F, 76.0F, 17.0F, False, PptHexColor("#DDE7F0", "#DDE7F0"), fontName, 2, 0.0F)
                End If

            Case Else
                addStandardTitle()
                Dim bulletItems As JArray = GetAutoPilotPowerPointBulletItems(slideObj)
                If templateLayoutApplied AndAlso bulletItems IsNot Nothing AndAlso bulletItems.Count > 0 AndAlso
                   TrySetPowerPointBestNarrativePlaceholderBulletItems(sld, bulletItems, New Integer() {2, 7, 17, 6}, 0) Then
                ElseIf templateLayoutApplied AndAlso TrySetPowerPointTemplateContent(sld, CleanPptBulletText(body), 0) Then
                ElseIf bulletItems IsNot Nothing AndAlso bulletItems.Count > 0 Then
                    RenderPptBulletItems(sld, bulletItems, slideW, fontName, textColor, light, lineColor, accent)
                Else
                    RenderPptBullets(sld, body, slideW, fontName, textColor, light, lineColor, accent)
                End If
        End Select

        If layout <> "title" AndAlso layout <> "section" AndAlso layout <> "closing" Then
            Dim footerText As String = GetArgString(args, "footer_text")
            If Not String.IsNullOrWhiteSpace(sourceText) Then footerText = sourceText
            If templateLayoutApplied Then
                ApplyPowerPointTemplateFooter(sld, footerText, slideIndex, GetArgBool(args, "show_slide_numbers", True))
            Else
                AddPptFooter(sld, footerText, slideIndex, slideW, slideH, fontName, muted, GetArgBool(args, "show_slide_numbers", True))
            End If
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
            ' TextFrame and TextFrame2 can expose different effective margins in PowerPoint.
            ' Rich-fit validation measures TextFrame2, so keep both APIs synchronized. Otherwise
            ' perfectly valid editable visuals can be rejected because TextFrame2 still reports
            ' inherited/default margins that the renderer never intended.
            Dim textFrame2 As Object = Nothing
            Try
                textFrame2 = tb.TextFrame2
                textFrame2.MarginLeft = margin
                textFrame2.MarginRight = margin
                textFrame2.MarginTop = margin
                textFrame2.MarginBottom = margin
            Catch ex As System.Exception
            Finally
                If textFrame2 IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(textFrame2) : Catch : End Try
            End Try
            tb.TextFrame.TextRange.Text = If(text, "")
            tb.TextFrame.TextRange.Font.Name = fontName
            tb.TextFrame.TextRange.Font.Size = fontSize
            tb.TextFrame.TextRange.Font.Bold = If(bold, -1, 0)
            tb.TextFrame.TextRange.Font.Color.RGB = fontColor
            tb.TextFrame.TextRange.ParagraphFormat.Alignment = alignment
            ' Shapes.AddTextbox can inherit bullet formatting from the active PowerPoint
            ' theme/master. Rich visual labels, titles, numbers and connector glyphs are
            ' ordinary text by default; renderers that intentionally need bullets enable
            ' them explicitly after this helper returns.
            tb.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = 0
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
            t = System.Text.RegularExpressions.Regex.Replace(t, "^(?:[•·▪◦‣⁃∙●○■□◆◇►▸\-\*\+]\s*)+", "").TrimStart()
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

    Private Shared Sub RenderPptBulletItems(sld As Object,
                                                  bulletItems As JArray,
                                                  slideW As Single,
                                                  fontName As String,
                                                  textColor As Integer,
                                                  light As Integer,
                                                  lineColor As Integer,
                                                  accent As Integer)
        If bulletItems Is Nothing OrElse bulletItems.Count = 0 Then Exit Sub
        Dim card As Object = AddPptShape(sld, 5, 48.0F, 128.0F, slideW - 96.0F, 334.0F, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.75F)
        If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch ex As System.Exception : End Try

        Dim bodyText As String = BuildAutoPilotPowerPointBulletBody(bulletItems)
        Dim fontSize As Single = GetPptNarrativeFontSize(bodyText, 18.0F, 16.5F, 15.0F)
        Dim tb As Object = AddPptTextBox(sld, bodyText, 78.0F, 153.0F, slideW - 156.0F, 286.0F, fontSize, False, textColor, fontName, 1, 0.0F)
        If tb Is Nothing Then Exit Sub
        Try
            ApplyAutoPilotPowerPointBulletItemsToShape(tb, bulletItems)
            Try
                tb.TextFrame.TextRange.ParagraphFormat.SpaceWithin = 1.05F
            Catch ex As System.Exception
            End Try
        Finally
            Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(tb) : Catch ex As System.Exception : End Try
        End Try
    End Sub

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
            Dim badgeWidth As Single = 0.0F
            If Not String.IsNullOrWhiteSpace(badge) Then
                badgeWidth = Math.Min(Math.Max(54.0F, 18.0F + CSng(badge.Length) * 6.2F), Math.Max(54.0F, cardW * 0.42F))
                AddPptTextBox(sld, badge, x + 18.0F, y + 22.0F, badgeWidth, 28.0F, 10.5F, True, toneColor, fontName, 1, 0.0F)
            End If

            Dim titleLeft As Single = x + 18.0F
            If badgeWidth > 0.0F Then titleLeft = x + 18.0F + badgeWidth + 10.0F
            Dim titleWidth As Single = Math.Max(72.0F, x + cardW - 18.0F - titleLeft)
            Dim cardTitle As String = If(cardObj.Value(Of String)("title"), "")
            Dim titleHeight As Single = If(cardTitle.Length > 30, 58.0F, 48.0F)
            Dim titleSize As Single = GetPptNarrativeFontSize(cardTitle, 17.0F, 15.5F, 14.0F)
            AddPptTextBox(sld, cardTitle, titleLeft, y + 18.0F, titleWidth, titleHeight, titleSize, True, toneColor, fontName, 1, 0.0F)

            Dim body As String = GetPptArrayText(cardObj, "body")
            If String.IsNullOrWhiteSpace(body) Then body = GetPptArrayText(cardObj, "items")
            Dim bodySize As Single = GetPptNarrativeFontSize(body, If(rows = 1, 15.0F, 14.0F), 13.0F, 12.0F)
            Dim bodyTop As Single = y + 28.0F + titleHeight
            AddPptTextBox(sld, body, x + 18.0F, bodyTop, cardW - 36.0F, Math.Max(42.0F, y + cardH - bodyTop - 16.0F), bodySize, False, textColor, fontName, 1, 0.0F)
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

            Dim stepTitle As String = If(stepObj.Value(Of String)("title"), "")
            Dim stepTitleH As Single = If(stepTitle.Length > 22, 62.0F, 50.0F)
            Dim stepTitleSize As Single = GetPptNarrativeFontSize(stepTitle, 15.0F, 14.0F, 13.0F)
            AddPptTextBox(sld, stepTitle, x + 16.0F, top + 31.0F, cardW - 32.0F, stepTitleH, stepTitleSize, True, numColor, fontName, 2, 0.0F)
            Dim body As String = GetPptArrayText(stepObj, "body")
            If String.IsNullOrWhiteSpace(body) Then body = GetPptArrayText(stepObj, "detail")
            Dim bodyTop As Single = top + 39.0F + stepTitleH
            Dim bodySize As Single = GetPptNarrativeFontSize(body, 13.0F, 12.0F, 11.0F)
            AddPptTextBox(sld, body, x + 16.0F, bodyTop, cardW - 32.0F, Math.Max(48.0F, top + cardH - bodyTop - 18.0F), bodySize, False, textColor, fontName, 2, 0.0F)

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

        Dim hierarchyNodes As JArray = TryCast(structureObj("nodes"), JArray)
        If hierarchyNodes IsNot Nothing AndAlso hierarchyNodes.Count > 0 Then
            RenderPptHierarchyNodes(sld, hierarchyNodes, slideW, fontName, textColor, muted, light, lineColor, accent, secondary)
            Exit Sub
        End If

        Dim topObj As JObject = TryCast(structureObj("top"), JObject)
        If topObj Is Nothing Then topObj = TryCast(structureObj("parent"), JObject)
        Dim children As JArray = TryCast(structureObj("children"), JArray)
        If topObj Is Nothing Then Exit Sub

        Dim topW As Single = Math.Min(360.0F, slideW - 220.0F)
        Dim topLeft As Single = (slideW - topW) / 2.0F
        Dim topY As Single = 150.0F
        Dim topH As Single = 146.0F

        Dim topCard As Object = AddPptShape(sld, 5, topLeft, topY, topW, topH, accent, accent, 0.0F)
        If topCard IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(topCard) : Catch : End Try
        AddPptTextBox(sld, topObj.Value(Of String)("title"), topLeft + 18.0F, topY + 16.0F, topW - 36.0F, 42.0F, 17.0F, True, PptHexColor("#FFFFFF", "#FFFFFF"), fontName, 2, 0.0F)
        AddPptTextBox(sld, GetPptArrayText(topObj, "body"), topLeft + 18.0F, topY + 61.0F, topW - 36.0F, 70.0F, 12.0F, False, PptHexColor("#E6EEF6", "#E6EEF6"), fontName, 2, 0.0F)

        If children Is Nothing OrElse children.Count = 0 Then Exit Sub
        Dim count As Integer = Math.Min(4, children.Count)
        Dim gap As Single = 20.0F
        Dim totalW As Single = slideW - 120.0F
        Dim childW As Single = (totalW - gap * (count - 1)) / count
        Dim childY As Single = 326.0F
        Dim childH As Single = 144.0F
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
            AddPptTextBox(sld, childObj.Value(Of String)("title"), x + 14.0F, childY + 15.0F, childW - 28.0F, 48.0F, 15.0F, True, secondary, fontName, 2, 0.0F)
            AddPptTextBox(sld, GetPptArrayText(childObj, "body"), x + 14.0F, childY + 64.0F, childW - 28.0F, 66.0F, 11.5F, False, textColor, fontName, 2, 0.0F)
        Next
    End Sub

    Private Shared Sub RenderPptHierarchyNodes(sld As Object,
                                                       nodes As JArray,
                                                       slideW As Single,
                                                       fontName As String,
                                                       textColor As Integer,
                                                       muted As Integer,
                                                       light As Integer,
                                                       lineColor As Integer,
                                                       accent As Integer,
                                                       secondary As Integer)
        If nodes Is Nothing OrElse nodes.Count = 0 Then Exit Sub

        Dim byId As New System.Collections.Generic.Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)
        For Each nodeObj As JObject In nodes.OfType(Of JObject)()
            Dim nodeId As String = If(nodeObj.Value(Of String)("id"), "").Trim()
            If nodeId <> "" AndAlso Not byId.ContainsKey(nodeId) Then byId(nodeId) = nodeObj
        Next
        If byId.Count = 0 Then Exit Sub

        Dim levels As New System.Collections.Generic.Dictionary(Of Integer, System.Collections.Generic.List(Of JObject))()
        Dim nodeLevels As New System.Collections.Generic.Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
        Dim maxLevel As Integer = 0

        For Each pair As System.Collections.Generic.KeyValuePair(Of String, JObject) In byId
            Dim level As Integer = 0
            Dim cursor As JObject = pair.Value
            Dim visited As New System.Collections.Generic.HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
            While cursor IsNot Nothing AndAlso level < 3
                Dim parentId As String = If(cursor.Value(Of String)("parent_id"), "").Trim()
                If parentId = "" OrElse Not byId.ContainsKey(parentId) OrElse Not visited.Add(parentId) Then Exit While
                level += 1
                cursor = byId(parentId)
            End While
            nodeLevels(pair.Key) = level
            maxLevel = Math.Max(maxLevel, level)
            If Not levels.ContainsKey(level) Then levels(level) = New System.Collections.Generic.List(Of JObject)()
            levels(level).Add(pair.Value)
        Next

        Dim topY As Single = 145.0F
        Dim bottomY As Single = 468.0F
        Dim levelCount As Integer = maxLevel + 1
        Dim verticalGap As Single = If(levelCount <= 2, 34.0F, 26.0F)
        ' Compute node height from the available vertical canvas first. The previous
        ' levelStep calculation could be smaller than nodeH, which made adjacent native
        ' cards overlap and hid their connectors. This geometry guarantees a visible
        ' inter-level corridor while remaining design/template agnostic.
        Dim availableNodeHeight As Single = bottomY - topY - verticalGap * Math.Max(0, levelCount - 1)
        Dim nodeH As Single = Math.Min(112.0F, availableNodeHeight / Math.Max(1, levelCount))
        nodeH = Math.Max(62.0F, nodeH)
        Dim levelStep As Single = nodeH + verticalGap
        Dim totalW As Single = slideW - 120.0F
        Dim rects As New System.Collections.Generic.Dictionary(Of String, AutoPilotPowerPointInteropRect)(StringComparer.OrdinalIgnoreCase)

        For level As Integer = 0 To maxLevel
            If Not levels.ContainsKey(level) Then Continue For
            Dim levelNodes As System.Collections.Generic.List(Of JObject) = levels(level)
            Dim count As Integer = levelNodes.Count
            Dim gap As Single = 18.0F
            Dim nodeW As Single = Math.Min(280.0F, (totalW - gap * Math.Max(0, count - 1)) / Math.Max(1, count))
            Dim usedW As Single = nodeW * count + gap * Math.Max(0, count - 1)
            Dim startX As Single = (slideW - usedW) / 2.0F
            Dim y As Single = topY + level * levelStep
            For i As Integer = 0 To count - 1
                Dim nodeObj As JObject = levelNodes(i)
                Dim nodeId As String = If(nodeObj.Value(Of String)("id"), "").Trim()
                rects(nodeId) = New AutoPilotPowerPointInteropRect With {
                    .X = startX + i * (nodeW + gap),
                    .Y = y,
                    .W = nodeW,
                    .H = nodeH
                }
            Next
        Next

        ' Draw connectors first so native editable node cards remain visually on top.
        For Each pair As System.Collections.Generic.KeyValuePair(Of String, JObject) In byId
            Dim parentId As String = If(pair.Value.Value(Of String)("parent_id"), "").Trim()
            If parentId = "" OrElse Not rects.ContainsKey(parentId) OrElse Not rects.ContainsKey(pair.Key) Then Continue For
            Dim parentRect As AutoPilotPowerPointInteropRect = rects(parentId)
            Dim childRect As AutoPilotPowerPointInteropRect = rects(pair.Key)
            Dim connector As Object = Nothing
            Try
                connector = sld.Shapes.AddLine(
                    parentRect.X + parentRect.W / 2.0F,
                    parentRect.Y + parentRect.H,
                    childRect.X + childRect.W / 2.0F,
                    childRect.Y)
                connector.Line.ForeColor.RGB = lineColor
                connector.Line.Weight = 1.4F
                connector.Line.EndArrowheadStyle = 2
            Finally
                If connector IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(connector) : Catch ex As System.Exception : End Try
            End Try
        Next

        For Each pair As System.Collections.Generic.KeyValuePair(Of String, JObject) In byId
            If Not rects.ContainsKey(pair.Key) Then Continue For
            Dim nodeObj As JObject = pair.Value
            Dim rect As AutoPilotPowerPointInteropRect = rects(pair.Key)
            Dim level As Integer = nodeLevels(pair.Key)
            Dim fillColor As Integer = If(level = 0, accent, PptHexColor("#FFFFFF", "#FFFFFF"))
            Dim outlineColor As Integer = If(level = 0, accent, If(level Mod 2 = 0, secondary, lineColor))
            Dim labelColor As Integer = If(level = 0, PptHexColor("#FFFFFF", "#FFFFFF"), If(level Mod 2 = 0, accent, secondary))
            Dim detailColor As Integer = If(level = 0, PptHexColor("#E6EEF6", "#E6EEF6"), textColor)

            Dim card As Object = AddPptShape(sld, 5, rect.X, rect.Y, rect.W, rect.H, fillColor, outlineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch ex As System.Exception : End Try

            Dim label As String = If(nodeObj.Value(Of String)("label"), "")
            Dim detail As String = If(nodeObj.Value(Of String)("detail"), "")
            Dim innerH As Single = Math.Max(30.0F, rect.H - 18.0F)
            Dim labelH As Single = If(detail = "", innerH, Math.Max(30.0F, Math.Min(48.0F, innerH * 0.48F)))
            Dim labelFont As Single = If(rect.H < 78.0F, 12.5F, 14.5F)
            Dim detailFont As Single = If(rect.H < 78.0F, 10.0F, 11.5F)
            AddPptTextBox(sld, label, rect.X + 12.0F, rect.Y + 9.0F, rect.W - 24.0F, labelH, labelFont, True, labelColor, fontName, 2, 0.0F)
            If detail <> "" Then
                Dim detailY As Single = rect.Y + 9.0F + labelH
                AddPptTextBox(sld, detail, rect.X + 12.0F, detailY, rect.W - 24.0F, Math.Max(16.0F, rect.Y + rect.H - detailY - 7.0F), detailFont, False, detailColor, fontName, 2, 0.0F)
            End If
        Next
    End Sub

    Private Shared Sub RenderPptTimeline(sld As Object,
                                         events As JArray,
                                         slideW As Single,
                                         slideH As Single,
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
        ' Keep the timeline axis in a stable vertical band across arbitrary 16:9-style
        ' templates. The previous fixed 255 pt axis left too little room for upper
        ' event details and allowed TextFrame2 recovery to grow them into the axis.
        Dim lineY As Single = Math.Max(245.0F, Math.Min(slideH - 245.0F, slideH * 0.53F))
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
            ' Timeline event titles commonly wrap to two lines. Give them enough native PowerPoint
            ' text-frame height up front instead of rejecting the complete editable visual after
            ' TextFrame2 measurement. The alternating geometry still keeps every event clear of
            ' the axis and leaves a separate detail area.
            ' Reserve independent vertical zones. This keeps upper detail text clear of
            ' the axis even when PowerPoint measures an extra wrapped line. It also means
            ' generic text-fit recovery normally has nothing to resize or reposition.
            Dim labelY As Single = If(above, lineY - 160.0F, lineY + 25.0F)
            Dim titleY As Single = If(above, lineY - 116.0F, lineY + 69.0F)
            Dim bodyY As Single = If(above, lineY - 67.0F, lineY + 118.0F)
            Dim labelH As Single = 40.0F
            Dim titleH As Single = 44.0F
            Dim bodyH As Single = If(above, 56.0F, 58.0F)
            Dim boxW As Single = Math.Min(150.0F, Math.Max(110.0F, stepW + 18.0F))
            AddPptTextBox(sld, eventObj.Value(Of String)("label"), x - boxW / 2.0F, labelY, boxW, labelH, 11.0F, True, colorValue, fontName, 2, 0.0F)
            AddPptTextBox(sld, eventObj.Value(Of String)("title"), x - boxW / 2.0F, titleY, boxW, titleH, 13.0F, True, textColor, fontName, 2, 0.0F)
            AddPptTextBox(sld, GetPptArrayText(eventObj, "body"), x - boxW / 2.0F, bodyY, boxW, bodyH, 10.5F, False, muted, fontName, 2, 0.0F)
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
        Dim top As Single = 158.0F
        Dim totalW As Single = slideW - 96.0F
        Dim colW As Single = (totalW - gap * (count - 1)) / count
        Dim h As Single = 310.0F

        For i As Integer = 0 To count - 1
            Dim colObj As JObject = TryCast(columns(i), JObject)
            If colObj Is Nothing Then Continue For
            Dim x As Single = left + i * (colW + gap)
            Dim toneColor As Integer = GetPptToneColor(colObj.Value(Of String)("tone"), If(i = 0, accent, secondary), secondary, muted)
            Dim card As Object = AddPptShape(sld, 5, x, top, colW, h, PptHexColor("#FFFFFF", "#FFFFFF"), lineColor, 0.8F)
            If card IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(card) : Catch : End Try
            AddPptShape(sld, 1, x, top, colW, 7.0F, toneColor, toneColor, 0.0F)
            Dim columnTitle As String = If(colObj.Value(Of String)("title"), "")
            Dim titleFontSize As Single = GetPptNarrativeFontSize(columnTitle, 16.0F, 15.0F, 14.0F)
            Dim titleHeight As Single = If(columnTitle.Length > 42, 62.0F, 50.0F)
            AddPptTextBox(sld, columnTitle, x + 18.0F, top + 18.0F, colW - 36.0F, titleHeight, titleFontSize, True, toneColor, fontName, 1, 0.0F)

            Dim items As String = GetPptArrayText(colObj, "items")
            Dim verdict As String = colObj.Value(Of String)("verdict")
            Dim verdictReserve As Single = If(String.IsNullOrWhiteSpace(verdict), 0.0F, 48.0F)
            Dim itemsTop As Single = top + 24.0F + titleHeight
            Dim itemHeight As Single = Math.Max(108.0F, top + h - itemsTop - 18.0F - verdictReserve)
            Dim itemFontSize As Single = GetPptNarrativeFontSize(items, 13.5F, 12.25F, 11.5F)
            Dim itemBox As Object = AddPptTextBox(sld, CleanPptBulletText(items), x + 18.0F, itemsTop, colW - 36.0F, itemHeight, itemFontSize, False, textColor, fontName, 1, 0.0F)
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

            Dim sourceFormatRoutingError As System.String = GetSourceFormatCreatorRoutingError(toolCall.Arguments, "Excel", context)
            If Not System.String.IsNullOrWhiteSpace(sourceFormatRoutingError) Then
                response.Success = False
                response.ErrorMessage = sourceFormatRoutingError
                response.Response = sourceFormatRoutingError
                If context IsNot Nothing Then context.Log(sourceFormatRoutingError, "warn")
                Return response
            End If

            If context IsNot Nothing AndAlso context.SequencingState IsNot Nothing Then
                SharedLibrary.Agents.ToolCallSequencing.CaptureRetryInvariantArguments(toolCall.ToolName, toolCall.Arguments, context.SequencingState)
            End If

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

    Private Shared Function GetAutoPilotWordFootnotes(args As Dictionary(Of String, Object)) As Newtonsoft.Json.Linq.JArray
        If args Is Nothing OrElse Not args.ContainsKey("footnotes") OrElse args("footnotes") Is Nothing Then
            Return New Newtonsoft.Json.Linq.JArray()
        End If

        Try
            Dim token As Newtonsoft.Json.Linq.JToken = TryCast(args("footnotes"), Newtonsoft.Json.Linq.JToken)
            If token Is Nothing Then
                token = Newtonsoft.Json.Linq.JToken.FromObject(args("footnotes"))
            End If
            If token IsNot Nothing AndAlso token.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                Return DirectCast(token, Newtonsoft.Json.Linq.JArray)
            End If
        Catch
        End Try

        Return New Newtonsoft.Json.Linq.JArray()
    End Function

    Private Shared Function ValidateAutoPilotWordFootnoteContract(
            markdownContent As System.String,
            footnotes As Newtonsoft.Json.Linq.JArray,
            ByRef validationError As System.String) As System.Boolean

        validationError = System.String.Empty
        If footnotes Is Nothing Then footnotes = New Newtonsoft.Json.Linq.JArray()

        Dim ids As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each token As Newtonsoft.Json.Linq.JToken In footnotes
            If token Is Nothing OrElse token.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then
                validationError = "Every create_word_document footnotes entry must be an object."
                Return False
            End If

            Dim footnote As Newtonsoft.Json.Linq.JObject = DirectCast(token, Newtonsoft.Json.Linq.JObject)
            Dim id As System.String = If(CStr(footnote("id")), System.String.Empty).Trim()
            Dim text As System.String = If(CStr(footnote("text")), System.String.Empty)
            If System.String.IsNullOrWhiteSpace(id) OrElse
               Not System.Text.RegularExpressions.Regex.IsMatch(id, "^[A-Za-z0-9_.-]{1,64}$", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                validationError = "Every create_word_document footnote requires a valid id."
                Return False
            End If
            If System.String.IsNullOrWhiteSpace(text) Then
                validationError = "Footnote '" & id & "' requires non-empty text."
                Return False
            End If
            If Not ids.Add(id) Then
                validationError = "Duplicate create_word_document footnote id '" & id & "' is not allowed."
                Return False
            End If

            Dim placeholder As System.String = "[[footnote:" & id & "]]"
            Dim placeholderCount As System.Int32 = CountOrdinalOccurrences(If(markdownContent, System.String.Empty), placeholder)
            If placeholderCount <> 1 Then
                validationError = "Footnote '" & id & "' requires exactly one " & placeholder & " marker in markdown_content; found " & placeholderCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                Return False
            End If
        Next

        Dim placeholderMatches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(
            If(markdownContent, System.String.Empty),
            "\[\[footnote:([A-Za-z0-9_.-]{1,64})\]\]",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        For Each placeholderMatch As System.Text.RegularExpressions.Match In placeholderMatches
            Dim placeholderId As System.String = placeholderMatch.Groups(1).Value
            If Not ids.Contains(placeholderId) Then
                validationError = "markdown_content contains [[footnote:" & placeholderId & "]] but no matching footnotes entry."
                Return False
            End If
        Next

        Return True
    End Function

    Private Shared Function ValidateAutoPilotWordCrossReferenceContract(
            markdownContent As System.String,
            ByRef referenceCount As System.Int32,
            ByRef validationError As System.String) As System.Boolean

        referenceCount = 0
        validationError = System.String.Empty
        Dim content As System.String = If(markdownContent, System.String.Empty).Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
        Dim anchorRegex As New System.Text.RegularExpressions.Regex(
            "(?m)^[ \t]*\[\[anchor:([A-Za-z0-9_.-]{1,64})\]\][ \t]*$",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        Dim anyAnchorRegex As New System.Text.RegularExpressions.Regex(
            "\[\[anchor:([A-Za-z0-9_.-]{1,64})\]\]",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        Dim referenceRegex As New System.Text.RegularExpressions.Regex(
            "\[\[ref:([A-Za-z0-9_.-]{1,64}):(number|text|full)\]\]",
            System.Text.RegularExpressions.RegexOptions.CultureInvariant Or System.Text.RegularExpressions.RegexOptions.IgnoreCase)

        Dim anchorMatches As System.Text.RegularExpressions.MatchCollection = anchorRegex.Matches(content)
        Dim allAnchorMatches As System.Text.RegularExpressions.MatchCollection = anyAnchorRegex.Matches(content)
        If allAnchorMatches.Count <> anchorMatches.Count Then
            validationError = "Every [[anchor:ID]] marker must appear on its own Markdown line immediately before the target paragraph or heading."
            Return False
        End If

        Dim anchors As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
        For Each anchorMatch As System.Text.RegularExpressions.Match In anchorMatches
            Dim anchorId As System.String = anchorMatch.Groups(1).Value
            If Not anchors.Add(anchorId) Then
                validationError = "Duplicate create_word_document cross-reference anchor '" & anchorId & "' is not allowed."
                Return False
            End If
        Next

        Dim referenceMatches As System.Text.RegularExpressions.MatchCollection = referenceRegex.Matches(content)
        referenceCount = referenceMatches.Count
        For Each referenceMatch As System.Text.RegularExpressions.Match In referenceMatches
            Dim anchorId As System.String = referenceMatch.Groups(1).Value
            If Not anchors.Contains(anchorId) Then
                validationError = "Cross-reference marker '" & referenceMatch.Value & "' requires a matching [[anchor:" & anchorId & "]] marker in markdown_content."
                Return False
            End If
        Next

        If content.IndexOf("[[ref:", System.StringComparison.OrdinalIgnoreCase) >= 0 AndAlso referenceMatches.Count = 0 Then
            validationError = "Invalid Word cross-reference syntax. Use [[ref:ID:number]], [[ref:ID:text]], or [[ref:ID:full]]."
            Return False
        End If

        Return True
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

    Private Shared Function GetAutoPilotWordTemplateFields(args As Dictionary(Of String, Object)) As Newtonsoft.Json.Linq.JObject
        If args Is Nothing OrElse Not args.ContainsKey("template_fields") OrElse args("template_fields") Is Nothing Then
            Return New Newtonsoft.Json.Linq.JObject()
        End If

        Try
            Dim token As Newtonsoft.Json.Linq.JToken = TryCast(args("template_fields"), Newtonsoft.Json.Linq.JToken)
            If token Is Nothing Then token = Newtonsoft.Json.Linq.JToken.FromObject(args("template_fields"))
            If token IsNot Nothing AndAlso token.Type = Newtonsoft.Json.Linq.JTokenType.Object Then
                Return DirectCast(token, Newtonsoft.Json.Linq.JObject)
            End If
        Catch ex As System.Exception
        End Try

        Return New Newtonsoft.Json.Linq.JObject()
    End Function

    Private Shared Function GetAutoPilotWordTemplateFieldToken(
            templateFields As Newtonsoft.Json.Linq.JObject,
            key As String) As Newtonsoft.Json.Linq.JToken

        If templateFields Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return Nothing
        For Each propertyItem As Newtonsoft.Json.Linq.JProperty In templateFields.Properties()
            If System.String.Equals(propertyItem.Name, key, System.StringComparison.OrdinalIgnoreCase) Then
                Return propertyItem.Value
            End If
        Next
        Return Nothing
    End Function

    Private Shared Function ValidateAutoPilotWordTemplateContractInputs(
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            templateFields As Newtonsoft.Json.Linq.JObject,
            markdownContent As String,
            ByRef validationError As String) As Boolean

        validationError = String.Empty
        If contract Is Nothing OrElse Not contract.HasSlots Then Return True
        If templateFields Is Nothing Then templateFields = New Newtonsoft.Json.Linq.JObject()

        Dim allowedFieldKeys As New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
        For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
            If slot Is Nothing Then Continue For

            If slot.UsesMarkdownContent Then
                If slot.Required AndAlso System.String.IsNullOrWhiteSpace(markdownContent) Then
                    validationError = "The selected Word template requires markdown_content for placeholder " & slot.Placeholder & "."
                    Return False
                End If
                Continue For
            End If

            Dim key As String = slot.TemplateFieldKey
            If key = "" Then Continue For
            allowedFieldKeys.Add(key)

            If slot.Required Then
                Dim fieldToken As Newtonsoft.Json.Linq.JToken = GetAutoPilotWordTemplateFieldToken(templateFields, key)
                Dim fieldValue As String = If(fieldToken Is Nothing OrElse fieldToken.Type = Newtonsoft.Json.Linq.JTokenType.Null, "", fieldToken.ToString())
                If System.String.IsNullOrWhiteSpace(fieldValue) Then
                    validationError = "The selected Word template requires template_fields.'" & key & "' for placeholder " & slot.Placeholder & "."
                    Return False
                End If
            End If
        Next

        For Each propertyItem As Newtonsoft.Json.Linq.JProperty In templateFields.Properties()
            If Not allowedFieldKeys.Contains(propertyItem.Name) Then
                validationError = "Unknown template_fields key '" & propertyItem.Name & "' for the selected Word design. Use only the keys exposed by its design guidance."
                Return False
            End If
        Next

        Return True
    End Function

    Private Shared Function GetAutoPilotWordTemplateSlotValue(
            slot As SharedLibrary.Agents.WordTemplateSlotDefinition,
            templateFields As Newtonsoft.Json.Linq.JObject,
            markdownContent As String) As String

        If slot Is Nothing Then Return String.Empty
        If slot.UsesMarkdownContent Then Return If(markdownContent, String.Empty)

        Dim key As String = slot.TemplateFieldKey
        If key = "" OrElse templateFields Is Nothing Then Return String.Empty
        Dim token As Newtonsoft.Json.Linq.JToken = GetAutoPilotWordTemplateFieldToken(templateFields, key)
        If token Is Nothing OrElse token.Type = Newtonsoft.Json.Linq.JTokenType.Null Then Return String.Empty
        Return token.ToString()
    End Function

    Private Shared Function FindAutoPilotWordTemplatePlaceholderPositions(
            doc As Microsoft.Office.Interop.Word.Document,
            placeholder As String) As System.Collections.Generic.List(Of System.Tuple(Of Integer, Integer))

        Dim result As New System.Collections.Generic.List(Of System.Tuple(Of Integer, Integer))()
        If doc Is Nothing OrElse System.String.IsNullOrWhiteSpace(placeholder) Then Return result

        Dim searchRange As Microsoft.Office.Interop.Word.Range = Nothing
        Dim finder As Microsoft.Office.Interop.Word.Find = Nothing
        Try
            searchRange = doc.Content.Duplicate
            Dim documentEnd As Integer = searchRange.End

            Do While searchRange.Start < documentEnd
                If finder IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                    finder = Nothing
                End If

                finder = searchRange.Find
                finder.ClearFormatting()
                finder.Text = placeholder
                finder.Forward = True
                finder.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                finder.MatchCase = False
                finder.MatchWildcards = False

                If Not finder.Execute() Then Exit Do

                Dim foundStart As Integer = searchRange.Start
                Dim foundEnd As Integer = searchRange.End
                result.Add(System.Tuple.Create(foundStart, foundEnd))

                If foundEnd >= documentEnd Then Exit Do
                searchRange.SetRange(foundEnd, documentEnd)
            Loop
        Catch ex As System.Exception
            result.Clear()
        Finally
            If finder IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
            If searchRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(searchRange) : Catch ex As System.Exception : End Try
        End Try

        Return result
    End Function

    Private Shared Function CountAutoPilotWordTemplatePlaceholderOccurrencesAllStories(
            doc As Microsoft.Office.Interop.Word.Document,
            placeholder As String) As Integer

        If doc Is Nothing OrElse System.String.IsNullOrWhiteSpace(placeholder) Then Return 0

        Dim total As Integer = 0
        Try
            For Each firstStory As Microsoft.Office.Interop.Word.Range In doc.StoryRanges
                Dim currentStory As Microsoft.Office.Interop.Word.Range = firstStory
                Do While currentStory IsNot Nothing
                    Dim searchRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Dim finder As Microsoft.Office.Interop.Word.Find = Nothing
                    Try
                        searchRange = currentStory.Duplicate
                        Dim storyEnd As Integer = searchRange.End

                        Do While searchRange.Start < storyEnd
                            If finder IsNot Nothing Then
                                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                                finder = Nothing
                            End If

                            finder = searchRange.Find
                            finder.ClearFormatting()
                            finder.Text = placeholder
                            finder.Forward = True
                            finder.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                            finder.MatchCase = False
                            finder.MatchWildcards = False

                            If Not finder.Execute() Then Exit Do
                            total += 1

                            Dim foundEnd As Integer = searchRange.End
                            If foundEnd >= storyEnd Then Exit Do
                            searchRange.SetRange(foundEnd, storyEnd)
                        Loop
                    Finally
                        If finder IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                        If searchRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(searchRange) : Catch ex As System.Exception : End Try
                    End Try

                    currentStory = currentStory.NextStoryRange
                Loop
            Next
        Catch ex As System.Exception
            Return -1
        End Try

        Return total
    End Function

    Private Shared Function ReplaceAutoPilotWordTemplateTextPlaceholderAllStories(
            doc As Microsoft.Office.Interop.Word.Document,
            placeholder As String,
            value As String,
            ByRef replacedCount As Integer,
            ByRef replacementError As String) As Boolean

        replacedCount = 0
        replacementError = String.Empty
        If doc Is Nothing OrElse System.String.IsNullOrWhiteSpace(placeholder) Then Return True

        Try
            For Each firstStory As Microsoft.Office.Interop.Word.Range In doc.StoryRanges
                Dim currentStory As Microsoft.Office.Interop.Word.Range = firstStory
                Do While currentStory IsNot Nothing
                    Dim searchRange As Microsoft.Office.Interop.Word.Range = Nothing
                    Dim finder As Microsoft.Office.Interop.Word.Find = Nothing
                    Try
                        searchRange = currentStory.Duplicate

                        Do While searchRange.Start < currentStory.End
                            If finder IsNot Nothing Then
                                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                                finder = Nothing
                            End If

                            finder = searchRange.Find
                            finder.ClearFormatting()
                            finder.Text = placeholder
                            finder.Forward = True
                            finder.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindStop
                            finder.MatchCase = False
                            finder.MatchWildcards = False

                            If Not finder.Execute() Then Exit Do

                            searchRange.Text = If(value, String.Empty)
                            replacedCount += 1

                            Dim nextStart As Integer = searchRange.End
                            Dim storyEnd As Integer = currentStory.End
                            If nextStart >= storyEnd Then Exit Do
                            searchRange.SetRange(nextStart, storyEnd)
                        Loop
                    Finally
                        If finder IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(finder) : Catch ex As System.Exception : End Try
                        If searchRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(searchRange) : Catch ex As System.Exception : End Try
                    End Try

                    currentStory = currentStory.NextStoryRange
                Loop
            Next
            Return True
        Catch ex As System.Exception
            replacementError = ex.Message
            Return False
        End Try
    End Function

    Private Shared Function NormalizeAutoPilotWordTemplateParagraphText(value As String) As String
        If value Is Nothing Then Return String.Empty
        Return value.Replace(vbCr, String.Empty).
                     Replace(vbLf, String.Empty).
                     Replace(ChrW(7), String.Empty).
                     Trim()
    End Function

    Private Shared Function AutoPilotWordDocumentContainsTemplateMarkers(
            doc As Microsoft.Office.Interop.Word.Document) As Boolean

        If doc Is Nothing Then Return False
        Try
            For Each firstStory As Microsoft.Office.Interop.Word.Range In doc.StoryRanges
                Dim currentStory As Microsoft.Office.Interop.Word.Range = firstStory
                Do While currentStory IsNot Nothing
                    Dim storyText As String = If(currentStory.Text, String.Empty)
                    If System.Text.RegularExpressions.Regex.IsMatch(
                        storyText,
                        "\[\[RI:[\p{L}\p{N}_.-]{1,64}\]\]",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then

                        Return True
                    End If
                    currentStory = currentStory.NextStoryRange
                Loop
            Next
            Return False
        Catch ex As System.Exception
            ' A failed full-story scan is treated conservatively. Structured output must not
            ' be accepted when unresolved marker state cannot be established deterministically.
            Return True
        End Try
    End Function

    Private Shared Function ValidateAutoPilotWordTemplateBodyStyles(
            doc As Microsoft.Office.Interop.Word.Document,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            ByRef validationError As String) As Boolean

        validationError = System.String.Empty
        If contract Is Nothing OrElse Not contract.HasBodyStyles Then Return True
        If doc Is Nothing Then
            validationError = "The Word body-style contract could not be validated because no document is open."
            Return False
        End If

        For Each definition As SharedLibrary.Agents.WordTemplateBodyStyleDefinition In contract.BodyStyles
            If definition Is Nothing Then Continue For
            Dim styleObject As Microsoft.Office.Interop.Word.Style = Nothing
            Try
                styleObject = doc.Styles.Item(definition.StyleName)
                If styleObject Is Nothing Then
                    validationError = "Word style '" & definition.StyleName & "' declared for semantic '" & definition.Semantic & "' was not found in the selected template."
                    Return False
                End If
            Catch ex As System.Exception
                validationError = "Word style '" & definition.StyleName & "' declared for semantic '" & definition.Semantic & "' was not found in the selected template."
                Return False
            Finally
                If styleObject IsNot Nothing Then
                    Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(styleObject) : Catch ex As System.Exception : End Try
                End If
            End Try
        Next

        Return True
    End Function

    Private Shared Function GetAutoPilotWordParagraphSemantic(
            paragraph As Microsoft.Office.Interop.Word.Paragraph) As String

        If paragraph Is Nothing Then Return System.String.Empty
        Try
            If paragraph.Range.Information(Microsoft.Office.Interop.Word.WdInformation.wdWithInTable) Then Return System.String.Empty
        Catch ex As System.Exception
        End Try

        Try
            Select Case paragraph.OutlineLevel
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel1 : Return "heading1"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel2 : Return "heading2"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel3 : Return "heading3"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel4 : Return "heading4"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel5 : Return "heading5"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel6 : Return "heading6"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel7 : Return "heading7"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel8 : Return "heading8"
                Case Microsoft.Office.Interop.Word.WdOutlineLevel.wdOutlineLevel9 : Return "heading9"
            End Select
        Catch ex As System.Exception
        End Try

        Try
            Dim listType As Microsoft.Office.Interop.Word.WdListType = paragraph.Range.ListFormat.ListType
            If listType <> Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then
                Dim level As Integer = 1
                Try
                    level = paragraph.Range.ListFormat.ListLevelNumber
                    If level < 1 Then level = 1
                Catch ex As System.Exception
                    level = 1
                End Try

                If listType = Microsoft.Office.Interop.Word.WdListType.wdListBullet Then
                    Return "bullet" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End If
                Return "numbered" & level.ToString(System.Globalization.CultureInfo.InvariantCulture)
            End If
        Catch ex As System.Exception
        End Try

        Return "paragraph"
    End Function

    Private Shared Function TryApplyAutoPilotWordBodyStyles(
            doc As Microsoft.Office.Interop.Word.Document,
            insertedStart As Integer,
            insertedEnd As Integer,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            ByRef stylingSummary As String,
            ByRef stylingError As String) As Boolean

        stylingSummary = System.String.Empty
        stylingError = System.String.Empty
        If contract Is Nothing OrElse Not contract.HasBodyStyles Then Return True
        If doc Is Nothing Then
            stylingError = "The Word body-style contract could not be applied because no document is open."
            Return False
        End If

        Dim styleMap As System.Collections.Generic.Dictionary(Of String, String) = contract.BuildNativeParagraphStyleMap()
        If styleMap.Count = 0 Then Return True
        If insertedEnd < insertedStart Then
            stylingError = "The inserted Word body range could not be determined for native style application."
            Return False
        End If

        Dim insertedRange As Microsoft.Office.Interop.Word.Range = Nothing
        Dim appliedCount As Integer = 0
        Try
            insertedRange = doc.Range(insertedStart, insertedEnd)
            For Each paragraph As Microsoft.Office.Interop.Word.Paragraph In insertedRange.Paragraphs
                Dim paragraphText As String = NormalizeAutoPilotWordTemplateParagraphText(paragraph.Range.Text)
                If paragraphText = System.String.Empty Then Continue For
                If System.Text.RegularExpressions.Regex.IsMatch(
                    paragraphText,
                    "^\[\[visual:[^\]]+\]\]$",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                    Continue For
                End If

                Dim semantic As String = GetAutoPilotWordParagraphSemantic(paragraph)
                If semantic = System.String.Empty Then Continue For
                If Not styleMap.ContainsKey(semantic) Then
                    ' Ordinary paragraphs may intentionally be left to the destination style when
                    ' no paragraph mapping is declared. Structural Markdown semantics are strict:
                    ' an unexpected heading/list semantic must never silently fall back.
                    If System.String.Equals(semantic, "paragraph", System.StringComparison.OrdinalIgnoreCase) Then Continue For
                    stylingError = "Word imported a body paragraph as structural semantic '" & semantic & "', but that semantic is not declared by the selected Word body-style contract. Generation was stopped to avoid a silent formatting fallback."
                    Return False
                End If

                Dim styleName As String = styleMap(semantic)
                Dim targetStyle As Microsoft.Office.Interop.Word.Style = Nothing
                Dim actualStyle As Microsoft.Office.Interop.Word.Style = Nothing
                Try
                    targetStyle = doc.Styles.Item(styleName)

                    ' Imported HTML list/heading numbering is direct formatting. Remove it before
                    ' assigning a native template style so only the template's numbering definition wins.
                    If semantic.StartsWith("heading", System.StringComparison.OrdinalIgnoreCase) OrElse
                       semantic.StartsWith("bullet", System.StringComparison.OrdinalIgnoreCase) OrElse
                       semantic.StartsWith("numbered", System.StringComparison.OrdinalIgnoreCase) Then
                        Try
                            If paragraph.Range.ListFormat.ListType <> Microsoft.Office.Interop.Word.WdListType.wdListNoNumbering Then
                                paragraph.Range.ListFormat.RemoveNumbers(Microsoft.Office.Interop.Word.WdNumberType.wdNumberParagraph)
                            End If
                        Catch ex As System.Exception
                        End Try
                    End If

                    paragraph.Range.Style = targetStyle
                    Try : paragraph.Range.ParagraphFormat.Reset() : Catch ex As System.Exception : End Try

                    actualStyle = TryCast(paragraph.Range.Style, Microsoft.Office.Interop.Word.Style)
                    If actualStyle Is Nothing OrElse
                       Not System.String.Equals(actualStyle.NameLocal, targetStyle.NameLocal, System.StringComparison.OrdinalIgnoreCase) Then
                        stylingError = "Word failed to apply native style '" & styleName & "' to semantic '" & semantic & "'."
                        Return False
                    End If
                    appliedCount += 1
                Catch ex As System.Exception
                    stylingError = "Failed to apply Word style '" & styleName & "' to semantic '" & semantic & "': " & ex.Message
                    Return False
                Finally
                    If actualStyle IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(actualStyle) : Catch ex As System.Exception : End Try
                    If targetStyle IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(targetStyle) : Catch ex As System.Exception : End Try
                End Try
            Next
        Catch ex As System.Exception
            stylingError = "Failed to apply the Word body-style contract: " & ex.Message
            Return False
        Finally
            If insertedRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(insertedRange) : Catch ex As System.Exception : End Try
        End Try

        stylingSummary = " Applied " & appliedCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " native body style assignment(s)."
        Return True
    End Function

    Private Shared Function TryInsertAutoPilotWordMarkdown(
            doc As Microsoft.Office.Interop.Word.Document,
            selection As Microsoft.Office.Interop.Word.Selection,
            markdownContent As String,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            ByRef stylingSummary As String,
            ByRef insertionError As String) As Boolean

        stylingSummary = System.String.Empty
        insertionError = System.String.Empty
        If doc Is Nothing OrElse selection Is Nothing Then
            insertionError = "Word Markdown insertion requires an open document and selection."
            Return False
        End If

        Dim startPosition As Integer = selection.Range.Start
        Try
            SharedMethods.InsertTextWithMarkdown(
                selection,
                markdownContent,
                TrailingCR:=False,
                PreserveDestinationParagraphFormatting:=(contract IsNot Nothing AndAlso contract.HasBodyStyles))

            Dim endPosition As Integer = selection.Range.End
            If contract IsNot Nothing AndAlso contract.HasBodyStyles Then
                If Not TryApplyAutoPilotWordBodyStyles(doc, startPosition, endPosition, contract, stylingSummary, insertionError) Then Return False
            End If
            Return True
        Catch ex As System.Exception
            insertionError = "Failed to insert Markdown into Word: " & ex.Message
            Return False
        End Try
    End Function

    Private Shared Function TryBindAutoPilotWordTemplateSlots(
            doc As Microsoft.Office.Interop.Word.Document,
            contract As SharedLibrary.Agents.WordTemplateBindingContract,
            templateFields As Newtonsoft.Json.Linq.JObject,
            markdownContent As String,
            ByRef bindingSummary As String,
            ByRef bindingError As String) As Boolean

        bindingSummary = String.Empty
        bindingError = String.Empty
        If contract Is Nothing OrElse Not contract.HasSlots Then Return True
        If doc Is Nothing Then
            bindingError = "The Word template binding contract could not be applied because no document was open."
            Return False
        End If
        If templateFields Is Nothing Then templateFields = New Newtonsoft.Json.Linq.JObject()

        ' Validate the physical template before changing anything. A guide/template mismatch
        ' is a hard error rather than a silent append-at-start fallback.
        For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
            If slot Is Nothing Then Continue For

            Dim allStoryCount As Integer = CountAutoPilotWordTemplatePlaceholderOccurrencesAllStories(doc, slot.Placeholder)
            If allStoryCount < 0 Then
                bindingError = "Word template placeholder locations could not be inspected deterministically for " & slot.Placeholder & "."
                Return False
            End If
            If allStoryCount = 0 Then
                bindingError = "Word template placeholder " & slot.Placeholder & " declared in '" &
                               System.IO.Path.GetFileName(contract.GuidancePath) & "' was not found in the template."
                Return False
            End If

            If System.String.Equals(slot.ContentMode, "markdown", System.StringComparison.OrdinalIgnoreCase) Then
                Dim mainStoryPositions As System.Collections.Generic.List(Of System.Tuple(Of Integer, Integer)) =
                    FindAutoPilotWordTemplatePlaceholderPositions(doc, slot.Placeholder)
                If allStoryCount <> 1 OrElse mainStoryPositions.Count <> 1 Then
                    bindingError = "Markdown Word template placeholder " & slot.Placeholder &
                                   " must occur exactly once in the main document story; found " &
                                   allStoryCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " occurrence(s) across all Word stories."
                    Return False
                End If
            End If
        Next

        Dim boundCount As Integer = 0

        ' Plain-text fields can live in the main body, tables, headers/footers, text frames,
        ' footnotes/endnotes, or another writable Word story. The marker name remains semantic-free;
        ' only the companion mapping decides which value is written there.
        For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
            If slot Is Nothing OrElse Not System.String.Equals(slot.ContentMode, "text", System.StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim value As String = GetAutoPilotWordTemplateSlotValue(slot, templateFields, markdownContent)
            Dim replacedForSlot As Integer = 0
            Dim replacementError As String = String.Empty
            If Not ReplaceAutoPilotWordTemplateTextPlaceholderAllStories(
                doc,
                slot.Placeholder,
                value,
                replacedForSlot,
                replacementError) Then

                bindingError = "Failed to fill Word template placeholder " & slot.Placeholder & ": " & replacementError
                Return False
            End If
            If replacedForSlot = 0 Then
                bindingError = "Word template placeholder " & slot.Placeholder & " disappeared before it could be filled."
                Return False
            End If
            boundCount += replacedForSlot
        Next

        ' Markdown slots must occupy a whole paragraph. Without a body-style contract, the
        ' existing placeholder supplies the legacy insertion formatting. With a body-style
        ' contract, destination formatting is preferred and the declared native paragraph styles
        ' are assigned deterministically after paste.
        For Each slot As SharedLibrary.Agents.WordTemplateSlotDefinition In contract.Slots
            If slot Is Nothing OrElse Not System.String.Equals(slot.ContentMode, "markdown", System.StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim positions As System.Collections.Generic.List(Of System.Tuple(Of Integer, Integer)) =
                FindAutoPilotWordTemplatePlaceholderPositions(doc, slot.Placeholder)
            If positions.Count <> 1 Then
                bindingError = "Markdown Word template placeholder " & slot.Placeholder & " could not be resolved uniquely."
                Return False
            End If

            Dim target As Microsoft.Office.Interop.Word.Range = Nothing
            Dim paragraphRange As Microsoft.Office.Interop.Word.Range = Nothing
            Try
                target = doc.Range(positions(0).Item1, positions(0).Item2)
                paragraphRange = target.Paragraphs(1).Range.Duplicate
                Dim paragraphText As String = NormalizeAutoPilotWordTemplateParagraphText(paragraphRange.Text)
                If Not System.String.Equals(paragraphText, slot.Placeholder, System.StringComparison.OrdinalIgnoreCase) Then
                    bindingError = "Markdown Word template placeholder " & slot.Placeholder &
                                   " must be the only visible content in its paragraph so the template's native paragraph formatting can be inherited safely."
                    Return False
                End If

                Dim value As String = GetAutoPilotWordTemplateSlotValue(slot, templateFields, markdownContent)
                If System.String.IsNullOrEmpty(value) Then
                    target.Text = String.Empty
                Else
                    target.Select()
                    Dim bodyStyleSummary As String = System.String.Empty
                    Dim bodyInsertError As String = System.String.Empty
                    If Not TryInsertAutoPilotWordMarkdown(
                        doc,
                        doc.Application.Selection,
                        value,
                        contract,
                        bodyStyleSummary,
                        bodyInsertError) Then

                        bindingError = bodyInsertError
                        Return False
                    End If
                    bindingSummary &= bodyStyleSummary
                End If
                boundCount += 1
            Catch ex As System.Exception
                bindingError = "Failed to fill Markdown Word template placeholder " & slot.Placeholder & ": " & ex.Message
                Return False
            Finally
                If paragraphRange IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(paragraphRange) : Catch ex As System.Exception : End Try
                If target IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(target) : Catch ex As System.Exception : End Try
            End Try
        Next

        If AutoPilotWordDocumentContainsTemplateMarkers(doc) Then
            bindingError = "The generated Word document still contains one or more unresolved [[RI:...]] template placeholders. Output was rejected."
            Return False
        End If

        bindingSummary = " Bound " & boundCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & " template placeholder occurrence(s) using " &
                         System.IO.Path.GetFileName(contract.GuidancePath) & "." & bindingSummary
        Return True
    End Function

    Private Shared Function IsAutoPilotWordTemplateUiCustomizationEntry(entryName As System.String) As Boolean
        Dim normalized As System.String = If(entryName, System.String.Empty).Replace("\", "/").TrimStart("/"c)
        If normalized = System.String.Empty Then Return False
        If System.String.Equals(normalized, "word/customizations.xml", System.StringComparison.OrdinalIgnoreCase) Then Return True
        If System.String.Equals(normalized, "word/_rels/customizations.xml.rels", System.StringComparison.OrdinalIgnoreCase) Then Return True
        If System.String.Equals(normalized, "word/attachedToolbars.bin", System.StringComparison.OrdinalIgnoreCase) Then Return True
        If normalized.StartsWith("customUI/", System.StringComparison.OrdinalIgnoreCase) Then Return True
        Return False
    End Function

    Private Shared Function IsAutoPilotWordTemplateUiCustomizationRelationshipType(value As System.String) As Boolean
        Dim normalized As System.String = If(value, System.String.Empty).Trim()
        If normalized = System.String.Empty Then Return False
        If normalized.IndexOf("keyMapCustomizations", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If normalized.IndexOf("attachedToolbars", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If normalized.IndexOf("ui/extensibility", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        If normalized.IndexOf("customUI", System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return True
        Return False
    End Function

    Private Shared Sub WriteAutoPilotWordTemplatePackageXml(
            inputEntry As System.IO.Compression.ZipArchiveEntry,
            outputEntry As System.IO.Compression.ZipArchiveEntry,
            mutate As System.Action(Of System.Xml.Linq.XDocument))

        Using inputStream As System.IO.Stream = inputEntry.Open()
            Dim xml As System.Xml.Linq.XDocument = System.Xml.Linq.XDocument.Load(inputStream, System.Xml.Linq.LoadOptions.PreserveWhitespace)
            If mutate IsNot Nothing Then mutate(xml)
            Using outputStream As System.IO.Stream = outputEntry.Open()
                xml.Save(outputStream, System.Xml.Linq.SaveOptions.DisableFormatting)
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' Materializes a slot-bound macro-free Word template as the actual DOCX package before Word
    ''' opens it. This intentionally avoids Documents.Add(Template:=...) for .dotx carriers: that
    ''' Word template-instantiation path can depend on template trust/macro storage even when the
    ''' document content itself is macro-free. The clone keeps document content/styles/layout and
    ''' removes template-only UI customizations (legacy key bindings/toolbars/custom UI), which do
    ''' not participate in document rendering or slot binding.
    ''' </summary>
    Private Shared Function TryMaterializeAutoPilotSlotBoundDotxAsDocx(
            templatePath As System.String,
            outputPath As System.String,
            ByRef preparationSummary As System.String,
            ByRef preparationError As System.String) As System.Boolean

        preparationSummary = System.String.Empty
        preparationError = System.String.Empty

        If System.String.IsNullOrWhiteSpace(templatePath) OrElse Not System.IO.File.Exists(templatePath) Then
            preparationError = "The selected slot-bound Word template carrier is unavailable."
            Return False
        End If
        If System.String.IsNullOrWhiteSpace(outputPath) Then
            preparationError = "No output path was available for the slot-bound Word template carrier."
            Return False
        End If
        If Not System.String.Equals(System.IO.Path.GetExtension(templatePath), ".dotx", System.StringComparison.OrdinalIgnoreCase) Then
            preparationError = "The safe slot-bound template materializer accepts only macro-free .dotx carriers."
            Return False
        End If

        Try
            If System.IO.File.Exists(outputPath) Then System.IO.File.Delete(outputPath)

            Using inputStream As New System.IO.FileStream(templatePath, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.ReadWrite)
                Using inputArchive As New System.IO.Compression.ZipArchive(inputStream, System.IO.Compression.ZipArchiveMode.Read, leaveOpen:=False)
                    Using outputStream As New System.IO.FileStream(outputPath, System.IO.FileMode.CreateNew, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                        Using outputArchive As New System.IO.Compression.ZipArchive(outputStream, System.IO.Compression.ZipArchiveMode.Create, leaveOpen:=False)
                            For Each inputEntry As System.IO.Compression.ZipArchiveEntry In inputArchive.Entries
                                Dim normalizedName As System.String = If(inputEntry.FullName, System.String.Empty).Replace("\", "/")
                                If IsAutoPilotWordTemplateUiCustomizationEntry(normalizedName) Then Continue For

                                Dim outputEntry As System.IO.Compression.ZipArchiveEntry = outputArchive.CreateEntry(
                                    inputEntry.FullName,
                                    System.IO.Compression.CompressionLevel.Optimal)

                                If normalizedName.EndsWith("/", System.StringComparison.Ordinal) Then Continue For

                                If System.String.Equals(normalizedName, "[Content_Types].xml", System.StringComparison.OrdinalIgnoreCase) Then
                                    WriteAutoPilotWordTemplatePackageXml(
                                        inputEntry,
                                        outputEntry,
                                        Sub(xml As System.Xml.Linq.XDocument)
                                            Dim root As System.Xml.Linq.XElement = xml.Root
                                            If root Is Nothing Then Throw New System.Exception("The Word template content-type manifest is empty.")

                                            For Each element As System.Xml.Linq.XElement In New System.Collections.Generic.List(Of System.Xml.Linq.XElement)(root.Elements())
                                                Dim contentTypeAttribute As System.Xml.Linq.XAttribute = element.Attribute("ContentType")
                                                Dim contentType As System.String = If(contentTypeAttribute Is Nothing, System.String.Empty, contentTypeAttribute.Value)

                                                If System.String.Equals(element.Name.LocalName, "Default", System.StringComparison.Ordinal) Then
                                                    If System.String.Equals(contentType, "application/vnd.ms-word.attachedToolbars", System.StringComparison.OrdinalIgnoreCase) Then
                                                        element.Remove()
                                                    End If
                                                    Continue For
                                                End If
                                                If Not System.String.Equals(element.Name.LocalName, "Override", System.StringComparison.Ordinal) Then Continue For

                                                Dim partNameAttribute As System.Xml.Linq.XAttribute = element.Attribute("PartName")
                                                Dim partName As System.String = If(partNameAttribute Is Nothing, System.String.Empty, partNameAttribute.Value)

                                                If System.String.Equals(partName, "/word/document.xml", System.StringComparison.OrdinalIgnoreCase) Then
                                                    If contentTypeAttribute Is Nothing Then
                                                        element.Add(New System.Xml.Linq.XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"))
                                                    Else
                                                        contentTypeAttribute.Value = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
                                                    End If
                                                ElseIf System.String.Equals(partName, "/word/customizations.xml", System.StringComparison.OrdinalIgnoreCase) OrElse
                                                       partName.StartsWith("/customUI/", System.StringComparison.OrdinalIgnoreCase) Then
                                                    element.Remove()
                                                End If
                                            Next
                                        End Sub)
                                    Continue For
                                End If

                                If System.String.Equals(normalizedName, "word/settings.xml", System.StringComparison.OrdinalIgnoreCase) Then
                                    WriteAutoPilotWordTemplatePackageXml(
                                        inputEntry,
                                        outputEntry,
                                        Sub(xml As System.Xml.Linq.XDocument)
                                            Dim root As System.Xml.Linq.XElement = xml.Root
                                            If root Is Nothing Then Return
                                            For Each element As System.Xml.Linq.XElement In New System.Collections.Generic.List(Of System.Xml.Linq.XElement)(root.Elements())
                                                If System.String.Equals(element.Name.LocalName, "attachedTemplate", System.StringComparison.Ordinal) Then element.Remove()
                                            Next
                                        End Sub)
                                    Continue For
                                End If

                                If System.String.Equals(normalizedName, "word/_rels/document.xml.rels", System.StringComparison.OrdinalIgnoreCase) OrElse
                                   System.String.Equals(normalizedName, "word/_rels/settings.xml.rels", System.StringComparison.OrdinalIgnoreCase) OrElse
                                   System.String.Equals(normalizedName, "_rels/.rels", System.StringComparison.OrdinalIgnoreCase) Then

                                    WriteAutoPilotWordTemplatePackageXml(
                                        inputEntry,
                                        outputEntry,
                                        Sub(xml As System.Xml.Linq.XDocument)
                                            Dim root As System.Xml.Linq.XElement = xml.Root
                                            If root Is Nothing Then Return
                                            For Each relationship As System.Xml.Linq.XElement In New System.Collections.Generic.List(Of System.Xml.Linq.XElement)(root.Elements())
                                                If Not System.String.Equals(relationship.Name.LocalName, "Relationship", System.StringComparison.Ordinal) Then Continue For
                                                Dim typeAttribute As System.Xml.Linq.XAttribute = relationship.Attribute("Type")
                                                Dim relationshipType As System.String = If(typeAttribute Is Nothing, System.String.Empty, typeAttribute.Value)
                                                If IsAutoPilotWordTemplateUiCustomizationRelationshipType(relationshipType) OrElse
                                                   relationshipType.EndsWith("/attachedTemplate", System.StringComparison.OrdinalIgnoreCase) Then
                                                    relationship.Remove()
                                                End If
                                            Next
                                        End Sub)
                                    Continue For
                                End If

                                Using entryInput As System.IO.Stream = inputEntry.Open()
                                    Using entryOutput As System.IO.Stream = outputEntry.Open()
                                        entryInput.CopyTo(entryOutput)
                                    End Using
                                End Using
                            Next
                        End Using
                    End Using
                End Using
            End Using

            If Not System.IO.File.Exists(outputPath) OrElse New System.IO.FileInfo(outputPath).Length = 0 Then
                preparationError = "The slot-bound Word template carrier could not be materialized as a DOCX package."
                Return False
            End If

            preparationSummary = " Materialized the macro-free .dotx carrier as a DOCX package before Word opened it; template-only UI customizations were excluded."
            Return True
        Catch ex As System.Exception
            Try
                If System.IO.File.Exists(outputPath) Then System.IO.File.Delete(outputPath)
            Catch cleanupEx As System.Exception
            End Try
            preparationError = "Failed to materialize the slot-bound .dotx carrier safely: " & ex.Message
            Return False
        End Try
    End Function

    Private Shared Function RefreshAutoPilotGeneratedWordFields(
            outputPath As System.String,
            ByRef updatedFieldCount As System.Int32,
            ByRef refreshError As System.String) As System.Boolean

        updatedFieldCount = 0
        refreshError = System.String.Empty
        If System.String.IsNullOrWhiteSpace(outputPath) OrElse Not System.IO.File.Exists(outputPath) Then
            refreshError = "Cannot refresh Word fields because the generated DOCX was not found."
            Return False
        End If

        Dim wordApp As Microsoft.Office.Interop.Word.Application = Nothing
        Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
        Try
            ' Cross-reference creation is normally OOXML-only. When native REF fields are
            ' present, use one isolated hidden Word instance at the very end so Word itself
            ' resolves its native heading/list numbering and writes the field-result cache.
            ' Never reuse the user's interactive Word instance and never persist updateFields.
            wordApp = New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = False
            wordApp.ScreenUpdating = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
            Try : wordApp.Options.UpdateLinksAtOpen = False : Catch ex As System.Exception : End Try

            doc = wordApp.Documents.Open(
                FileName:=outputPath,
                ConfirmConversions:=False,
                ReadOnly:=False,
                AddToRecentFiles:=False,
                Revert:=False,
                Visible:=False,
                OpenAndRepair:=False)

            Try : doc.Repaginate() : Catch ex As System.Exception : End Try

            ' Update every field story in this newly generated document in one bounded pass.
            ' This deliberately avoids per-reference Interop calls; 1 or 300 REF fields incur
            ' the same single open/update/save lifecycle. StoryTypes 1..17 cover main text,
            ' notes/comments, text frames, headers/footers and separator stories.
            For storyTypeValue As System.Int32 = 1 To 17
                Dim currentStory As Microsoft.Office.Interop.Word.Range = Nothing
                Try
                    currentStory = doc.StoryRanges(CType(storyTypeValue, Microsoft.Office.Interop.Word.WdStoryType))
                Catch ex As System.Exception
                    currentStory = Nothing
                End Try

                Do While currentStory IsNot Nothing
                    Dim fields As Microsoft.Office.Interop.Word.Fields = Nothing
                    Dim nextStory As Microsoft.Office.Interop.Word.Range = Nothing
                    Try
                        fields = currentStory.Fields
                        Dim count As System.Int32 = 0
                        Try : count = fields.Count : Catch ex As System.Exception : End Try
                        If count > 0 Then
                            fields.Update()
                            updatedFieldCount += count
                        End If
                        Try : nextStory = currentStory.NextStoryRange : Catch ex As System.Exception : nextStory = Nothing : End Try
                    Finally
                        If fields IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(fields) : Catch ex As System.Exception : End Try
                        If currentStory IsNot Nothing Then Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(currentStory) : Catch ex As System.Exception : End Try
                    End Try
                    currentStory = nextStory
                Loop
            Next

            doc.Save()
            Return True
        Catch ex As System.Exception
            refreshError = "Word field refresh failed: " & ex.Message
            Return False
        Finally
            If doc IsNot Nothing Then
                Try : doc.Close(SaveChanges:=False) : Catch ex As System.Exception : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(doc) : Catch ex As System.Exception : End Try
            End If
            If wordApp IsNot Nothing Then
                Try : wordApp.ScreenUpdating = True : Catch ex As System.Exception : End Try
                Try : wordApp.Quit(SaveChanges:=False) : Catch ex As System.Exception : End Try
                Try : System.Runtime.InteropServices.Marshal.FinalReleaseComObject(wordApp) : Catch ex As System.Exception : End Try
            End If
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

            Dim sourceFormatRoutingError As System.String = GetSourceFormatCreatorRoutingError(toolCall.Arguments, "Word", context)
            If Not System.String.IsNullOrWhiteSpace(sourceFormatRoutingError) Then
                response.Success = False
                response.ErrorMessage = sourceFormatRoutingError
                response.Response = sourceFormatRoutingError
                If context IsNot Nothing Then context.Log(sourceFormatRoutingError, "warn")
                Return response
            End If

            If context IsNot Nothing AndAlso context.SequencingState IsNot Nothing Then
                SharedLibrary.Agents.ToolCallSequencing.CaptureRetryInvariantArguments(toolCall.ToolName, toolCall.Arguments, context.SequencingState)
            End If

            If GetArgBool(toolCall.Arguments, "use_repository_default_design", True) AndAlso
               HasMeaningfulToolArgument(toolCall.Arguments, "document_type") AndAlso
               (design Is Nothing OrElse design.Descriptor Is Nothing) Then
                response.Success = False
                response.ErrorMessage = If(If(design Is Nothing, System.String.Empty, design.TemplateWarning), System.String.Empty)
                If System.String.IsNullOrWhiteSpace(response.ErrorMessage) Then response.ErrorMessage = "No configured Word design matched the requested document type. The host will not silently substitute a generic/blank design."
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim wordTemplateContract As SharedLibrary.Agents.WordTemplateBindingContract = Nothing
            Dim wordTemplateContractError As String = String.Empty
            If design IsNot Nothing AndAlso design.Descriptor IsNot Nothing AndAlso design.ApplicationConfig IsNot Nothing Then
                If Not SharedLibrary.Agents.WordTemplateBindingContractParser.TryLoadForDesign(
                    design.Descriptor,
                    design.ApplicationConfig,
                    design.TemplatePath,
                    wordTemplateContract,
                    wordTemplateContractError) Then

                    response.Success = False
                    response.ErrorMessage = wordTemplateContractError
                    response.Response = response.ErrorMessage
                    Return response
                End If
            End If

            Dim templateFields As Newtonsoft.Json.Linq.JObject = GetAutoPilotWordTemplateFields(toolCall.Arguments)
            If wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasSlots Then
                If String.IsNullOrWhiteSpace(design.TemplatePath) Then
                    response.Success = False
                    response.ErrorMessage = "The selected Word design declares structured template slots, but no usable Word template carrier is available."
                    response.Response = response.ErrorMessage
                    Return response
                End If

                Dim templateInputError As String = String.Empty
                If Not ValidateAutoPilotWordTemplateContractInputs(wordTemplateContract, templateFields, markdownContent, templateInputError) Then
                    response.Success = False
                    response.ErrorMessage = templateInputError
                    response.Response = response.ErrorMessage
                    Return response
                End If

                If GetArgBool(toolCall.Arguments, "include_cover", False) Then
                    ' Presentation-only options must not invalidate a structurally bound design.
                    ' The carrier owns its document structure, so include_cover is a tolerant no-op here.
                    If context IsNot Nothing Then context.Log("Ignored include_cover=true because the selected slot-bound Word design owns document structure.")
                End If

                If context IsNot Nothing Then
                    context.Log("Word template slot contract loaded: " & wordTemplateContract.BuildPromptSummary())
                End If
            End If

            If wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasBodyStyles Then
                Dim paragraphStyleMap As System.Collections.Generic.Dictionary(Of String, String) = wordTemplateContract.BuildNativeParagraphStyleMap()
                Dim markdownStyleError As String = System.String.Empty
                If Not SharedMethods.ValidateMarkdownParagraphStyleMap(markdownContent, paragraphStyleMap, markdownStyleError) Then
                    response.Success = False
                    response.ErrorMessage = markdownStyleError
                    response.Response = response.ErrorMessage
                    Return response
                End If

                If context IsNot Nothing Then
                    context.Log("Word body-style contract loaded: " & wordTemplateContract.BuildPromptSummary())
                End If
            End If

            Dim footnotes As Newtonsoft.Json.Linq.JArray = GetAutoPilotWordFootnotes(toolCall.Arguments)
            Dim footnoteContractError As System.String = System.String.Empty
            If Not ValidateAutoPilotWordFootnoteContract(markdownContent, footnotes, footnoteContractError) Then
                response.Success = False
                response.ErrorMessage = footnoteContractError
                response.Response = response.ErrorMessage
                Return response
            End If

            Dim requestedCrossReferenceCount As System.Int32 = 0
            Dim crossReferenceContractError As System.String = System.String.Empty
            If Not ValidateAutoPilotWordCrossReferenceContract(markdownContent, requestedCrossReferenceCount, crossReferenceContractError) Then
                response.Success = False
                response.ErrorMessage = crossReferenceContractError
                response.Response = response.ErrorMessage
                Return response
            End If

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
            Dim templateBindingSummary As String = String.Empty

            Dim success As System.Boolean
            If wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasSlots Then
                If context IsNot Nothing Then
                    If requestedCrossReferenceCount > 0 Then
                        context.Log("Structured Word template renderer selected: OOXML-first; one final hidden Word field-refresh pass will run for native cross-references.")
                    Else
                        context.Log("Structured Word template renderer selected: OOXML-only; Word/COM will not be started.")
                    End If
                End If
                success = TryCreateAutoPilotStructuredWordDocumentOpenXml(
                    design.TemplatePath,
                    outputPath,
                    wordTemplateContract,
                    templateFields,
                    markdownContent,
                    GetArgString(toolCall.Arguments, "table_style_name"),
                    templateBindingSummary,
                    creationError)
            ElseIf design Is Nothing OrElse System.String.IsNullOrWhiteSpace(design.TemplatePath) Then
                If context IsNot Nothing Then
                    If requestedCrossReferenceCount > 0 Then
                        context.Log("Generic Word renderer selected: OOXML-first; one final hidden Word field-refresh pass will run for native cross-references.")
                    Else
                        context.Log("Generic Word renderer selected: OOXML-only; Word/COM will not be started.")
                    End If
                End If
                success = TryCreateAutoPilotGenericWordDocumentOpenXml(
                    outputPath,
                    markdownContent,
                    toolCall.Arguments,
                    templateBindingSummary,
                    creationError)
            Else
                If context IsNot Nothing Then context.Log("Legacy Word carrier renderer selected: Word/COM compatibility path.")
                success = Await SwitchToUi(Function()
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
                                                       If designExt = ".dotx" AndAlso wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasSlots Then
                                                           Dim carrierPreparationSummary As System.String = System.String.Empty
                                                           Dim carrierPreparationError As System.String = System.String.Empty
                                                           If Not TryMaterializeAutoPilotSlotBoundDotxAsDocx(
                                                               design.TemplatePath,
                                                               outputPath,
                                                               carrierPreparationSummary,
                                                               carrierPreparationError) Then

                                                               Throw New System.Exception(carrierPreparationError)
                                                           End If
                                                           templateBindingSummary &= carrierPreparationSummary
                                                           doc = wordApp.Documents.Open(outputPath, ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)
                                                       ElseIf designExt = ".dotx" OrElse designExt = ".dotm" Then
                                                           doc = wordApp.Documents.Add(Template:=design.TemplatePath, NewTemplate:=False)
                                                       Else
                                                           ' A .docx design source is cloned before use. Legacy style carriers keep
                                                           ' the established clear-body behavior. A slot-bound template is a native
                                                           ' document structure carrier, so its body must remain intact for binding.
                                                           System.IO.File.Copy(design.TemplatePath, outputPath, overwrite:=False)
                                                           doc = wordApp.Documents.Open(outputPath, ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)
                                                           If wordTemplateContract Is Nothing OrElse Not wordTemplateContract.HasSlots Then
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
                                                       End If
                                                   Else
                                                       doc = wordApp.Documents.Add()
                                                   End If
                                                   doc.Activate()

                                                   sel = wordApp.Selection

                                                   If wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasBodyStyles Then
                                                       Dim bodyStyleValidationError As String = System.String.Empty
                                                       If Not ValidateAutoPilotWordTemplateBodyStyles(doc, wordTemplateContract, bodyStyleValidationError) Then
                                                           Throw New System.Exception(bodyStyleValidationError)
                                                       End If
                                                   End If

                                                   If wordTemplateContract IsNot Nothing AndAlso wordTemplateContract.HasSlots Then
                                                       Dim bindingError As String = String.Empty
                                                       If Not TryBindAutoPilotWordTemplateSlots(
                                                           doc,
                                                           wordTemplateContract,
                                                           templateFields,
                                                           markdownContent,
                                                           templateBindingSummary,
                                                           bindingError) Then

                                                           Throw New System.Exception(bindingError)
                                                       End If
                                                   Else
                                                       ' Preserve the legacy create-word path exactly for designs that have
                                                       ' not opted into the structured slot contract. If RI markers are present,
                                                       ' however, a missing companion mapping is unsafe and must not be guessed.
                                                       If design IsNot Nothing AndAlso
                                                          Not String.IsNullOrWhiteSpace(design.TemplatePath) AndAlso
                                                          AutoPilotWordDocumentContainsTemplateMarkers(doc) Then
                                                           Throw New System.Exception("The Word template contains [[RI:...]] placeholders but no valid Word template slot contract was loaded from its companion guidance file.")
                                                       End If

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

                                                       Dim legacyBodyStyleSummary As String = System.String.Empty
                                                       Dim legacyInsertError As String = System.String.Empty
                                                       If Not TryInsertAutoPilotWordMarkdown(
                                                           doc,
                                                           sel,
                                                           markdownContent,
                                                           wordTemplateContract,
                                                           legacyBodyStyleSummary,
                                                           legacyInsertError) Then

                                                           Throw New System.Exception(legacyInsertError)
                                                       End If
                                                       templateBindingSummary &= legacyBodyStyleSummary
                                                   End If

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
            End If

            Dim insertedFootnoteCount As System.Int32 = 0
            If success AndAlso File.Exists(outputPath) AndAlso footnotes.Count > 0 Then
                Dim footnoteInsertionError As System.String = System.String.Empty
                If Not InsertAutoPilotWordFootnotesOpenXml(outputPath, footnotes, insertedFootnoteCount, footnoteInsertionError) Then
                    success = False
                    creationError = footnoteInsertionError
                End If
            End If

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

            Dim insertedCrossReferenceAnchorCount As System.Int32 = 0
            Dim insertedCrossReferenceCount As System.Int32 = 0
            Dim refreshedFieldCount As System.Int32 = 0
            Dim crossReferenceSyntaxPresent As System.Boolean =
                requestedCrossReferenceCount > 0 OrElse
                markdownContent.IndexOf("[[anchor:", System.StringComparison.OrdinalIgnoreCase) >= 0

            If success AndAlso File.Exists(outputPath) AndAlso crossReferenceSyntaxPresent Then
                Dim crossReferenceInsertionError As System.String = System.String.Empty
                If Not InsertAutoPilotWordCrossReferencesOpenXml(
                    outputPath,
                    insertedCrossReferenceAnchorCount,
                    insertedCrossReferenceCount,
                    crossReferenceInsertionError) Then

                    success = False
                    creationError = crossReferenceInsertionError
                ElseIf insertedCrossReferenceCount <> requestedCrossReferenceCount Then
                    success = False
                    creationError = "Native Word cross-reference insertion count mismatch: expected " &
                                    requestedCrossReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                    ", inserted " & insertedCrossReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture) & "."
                End If
            End If

            If success AndAlso File.Exists(outputPath) AndAlso insertedCrossReferenceCount > 0 Then
                ct.ThrowIfCancellationRequested()
                If context IsNot Nothing Then context.Log("Native Word cross-references detected; running one final hidden Word field-refresh pass.")
                Dim fieldRefreshError As System.String = System.String.Empty
                Dim fieldRefreshSucceeded As System.Boolean = Await SwitchToUi(
                    Function() RefreshAutoPilotGeneratedWordFields(outputPath, refreshedFieldCount, fieldRefreshError))
                If Not fieldRefreshSucceeded Then
                    success = False
                    creationError = fieldRefreshError
                Else
                    Dim fieldStateError As System.String = System.String.Empty
                    If Not NormalizeAutoPilotWordOpenXmlFieldUpdateStateOnDisk(outputPath, fieldStateError) Then
                        success = False
                        creationError = fieldStateError
                    Else
                        Dim crossReferenceValidationError As System.String = System.String.Empty
                        If Not ValidateAutoPilotWordCrossReferenceRefreshOpenXml(outputPath, insertedCrossReferenceCount, crossReferenceValidationError) Then
                            success = False
                            creationError = crossReferenceValidationError
                        End If
                    End If
                End If
            End If

            If success AndAlso File.Exists(outputPath) Then
                RegisterAutoPilotGeneratedOutputFile(outputPath)

                response.Success = True
                Dim designSummary As String = BuildDesignExecutionNote(design)
                Dim footnoteSummary As System.String = System.String.Empty
                If footnotes.Count > 0 Then
                    footnoteSummary = " Inserted " & insertedFootnoteCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                      "/" & footnotes.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) & " native Word footnote(s)."
                End If
                Dim visualSummary As String = String.Empty
                If visuals.Count > 0 Then
                    visualSummary = $" Embedded {embeddedVisualCount}/{visuals.Count} requested visual(s)."
                    If visualWarnings.Count > 0 Then
                        visualSummary &= " Visual warnings: " & String.Join(" | ", visualWarnings)
                    End If
                End If
                Dim crossReferenceSummary As System.String = System.String.Empty
                If insertedCrossReferenceCount > 0 Then
                    crossReferenceSummary = " Inserted " & insertedCrossReferenceCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                            " native Word cross-reference(s) across " & insertedCrossReferenceAnchorCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                            " bookmark target(s), then refreshed " & refreshedFieldCount.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                                            " field(s) in one final Word pass."
                End If
                response.Response = $"Word document created: {fileName} ({New FileInfo(outputPath).Length / 1024:F0} KB). The file will be attached to the reply.{designSummary}{templateBindingSummary}{footnoteSummary}{visualSummary}{crossReferenceSummary}"
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

            Dim globalContext = GetArgString(toolCall.Arguments, "global_context")
            Dim processorInstruction As String = instruction.Trim()
            If Not String.IsNullOrWhiteSpace(globalContext) Then
                processorInstruction &= vbCrLf & vbCrLf &
                    "[DOCUMENT-WIDE CONSISTENCY CONTEXT - apply this guidance to every chunk; this is context, not a separate operation]" &
                    vbCrLf & globalContext.Trim() & vbCrLf &
                    "[/DOCUMENT-WIDE CONSISTENCY CONTEXT]"
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
                Dim success = Await ProcessDocumentForAutoPilot(inputPath, outputPath, processorInstruction, ct, sheetFilter, useOfflineDocs)

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
                            resultMessages.Add($"✓ {att.OriginalFileName}: Processed successfully. Created files: {outputName}; {Path.GetFileName(comparePath)} (tracked-changes compare).")
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

        Try
            ' Always use a dedicated hidden Word instance for comparison. Never attach to an
            ' existing user/host instance: the compare lifecycle is owned by this method and
            ' must be closed deterministically when the comparison finishes.
            wordApp = New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = False
            wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone
            wordApp.ScreenUpdating = False
            Try : wordApp.Options.UpdateLinksAtOpen = False : Catch : End Try

            originalDoc = wordApp.Documents.Open(
                FileName:=originalPath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

            processedDoc = wordApp.Documents.Open(
                FileName:=processedPath,
                ReadOnly:=True,
                Visible:=False,
                AddToRecentFiles:=False)

            compareDoc = wordApp.CompareDocuments(
                OriginalDocument:=originalDoc,
                RevisedDocument:=processedDoc,
                Destination:=Microsoft.Office.Interop.Word.WdCompareDestination.wdCompareDestinationNew,
                Granularity:=Microsoft.Office.Interop.Word.WdGranularity.wdGranularityWordLevel,
                CompareFormatting:=True,
                CompareCaseChanges:=True,
                CompareWhitespace:=True,
                CompareTables:=True,
                CompareHeaders:=True,
                CompareFootnotes:=True,
                CompareTextboxes:=True,
                CompareFields:=True,
                CompareComments:=True,
                RevisedAuthor:=AN6,
                IgnoreAllComparisonWarnings:=True)

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

            Return File.Exists(comparePath)

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
                Try : wordApp.Quit(Microsoft.Office.Interop.Word.WdSaveOptions.wdDoNotSaveChanges) : Catch : End Try
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
            Dim trackedChangesOnly = GetArgBool(toolCall.Arguments, "tracked_changes_only", False)
            If trackedChangesOnly Then includeTrackedChanges = True
            Dim filterAuthors As New List(Of String)()
            Dim legacyFilterAuthor = GetArgString(toolCall.Arguments, "tracked_changes_author")
            If Not String.IsNullOrWhiteSpace(legacyFilterAuthor) Then filterAuthors.Add(legacyFilterAuthor.Trim())
            For Each requestedAuthor As String In GetArgStringArray(toolCall.Arguments, "tracked_changes_authors")
                If Not String.IsNullOrWhiteSpace(requestedAuthor) AndAlso
                   Not filterAuthors.Any(Function(existingAuthor) String.Equals(existingAuthor, requestedAuthor.Trim(), StringComparison.OrdinalIgnoreCase)) Then
                    filterAuthors.Add(requestedAuthor.Trim())
                End If
            Next

            Dim filterSinceStr = GetArgString(toolCall.Arguments, "tracked_changes_since")
            Dim filterUntilStr = GetArgString(toolCall.Arguments, "tracked_changes_until")
            Dim filterSince As DateTime? = Nothing
            Dim filterUntil As DateTime? = Nothing
            Dim dateError As String = Nothing
            If Not TryParseRevisionDateFilter(filterSinceStr, False, filterSince, dateError) Then
                response.Success = False
                response.Response = dateError
                Return response
            End If
            If Not TryParseRevisionDateFilter(filterUntilStr, True, filterUntil, dateError) Then
                response.Success = False
                response.Response = dateError
                Return response
            End If
            If filterSince.HasValue AndAlso filterUntil.HasValue AndAlso filterSince.Value > filterUntil.Value Then
                response.Success = False
                response.Response = "tracked_changes_since must be earlier than or equal to tracked_changes_until."
                Return response
            End If

            context.Log($"Deep-reading Word document: {fileName}")
            ApDashboardLog($"📖 Deep-reading: {fileName}", "step")

            Dim result = Await System.Threading.Tasks.Task.Run(Function() ExtractWordDocumentDetails(
                att.TempFilePath, includeComments, includeHeadersFooters,
                includeFootnotesEndnotes, includeTrackedChanges, trackedChangesOnly, filterAuthors, filterSince, filterUntil))

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
            trackedChangesOnly As Boolean,
            filterAuthors As List(Of String),
            filterSince As DateTime?,
            filterUntil As DateTime?) As String

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

            ' Revision-only mode is intentionally a projection of the same XML reader, not a
            ' separate parser. This keeps author/date filtering and OOXML coverage identical while
            ' avoiding a very large body-text payload when the caller only needs tracked changes.
            If trackedChangesOnly Then
                AppendTrackedChangesXmlDetails(tempDir, sb, filterAuthors, filterSince, filterUntil)
                Return sb.ToString().TrimEnd()
            End If

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

                    ' Descendant selection intentionally includes paragraphs inside tables, content controls,
                    ' text boxes represented in the main document part, and other nested body structures.
                    For Each paraNode As XmlNode In bodyNode.SelectNodes(".//w:p", nsMgr)
                        Dim paraText As New StringBuilder()

                        For Each child As XmlNode In paraNode.ChildNodes
                            ProcessDocBodyNode(child, nsMgr, paraText, includeTrackedChanges,
                                             filterAuthors, filterSince, filterUntil,
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

                        AppendTrackedChangesXmlDetails(tempDir, sb, filterAuthors, filterSince, filterUntil)
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
            includeTrackedChanges As Boolean, filterAuthors As List(Of String), filterSince As DateTime?, filterUntil As DateTime?,
            ByRef insCount As Integer, ByRef delCount As Integer, ByRef fmtCount As Integer,
            authorCounts As Dictionary(Of String, Integer))

        If node Is Nothing Then Return

        Select Case node.LocalName
            Case "r" ' Normal run; preserve text while still inspecting run-property revisions.
                For Each child As XmlNode In node.ChildNodes
                    If child.LocalName = "t" Then
                        sb.Append(child.InnerText)
                    ElseIf child.LocalName = "tab" Then
                        sb.Append(vbTab)
                    ElseIf child.LocalName = "br" OrElse child.LocalName = "cr" Then
                        sb.Append(vbLf)
                    ElseIf child.LocalName = "rPr" Then
                        For Each propertyChild As XmlNode In child.ChildNodes
                            ProcessDocBodyNode(propertyChild, nsMgr, sb, includeTrackedChanges,
                                             filterAuthors, filterSince, filterUntil,
                                             insCount, delCount, fmtCount, authorCounts)
                        Next
                    End If
                Next

            Case "ins" ' Insertion
                Dim author = If(DirectCast(node, XmlElement).GetAttribute("w:author"), "")
                Dim dateStr = If(DirectCast(node, XmlElement).GetAttribute("w:date"), "")
                Dim shortDate = If(dateStr.Length >= 10, dateStr.Substring(0, 10), dateStr)

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthors, filterSince, filterUntil)

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

                Dim passesFilter = PassesRevisionFilter(author, dateStr, filterAuthors, filterSince, filterUntil)

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
                    If PassesRevisionFilter(author, dateStr, filterAuthors, filterSince, filterUntil) Then
                        fmtCount += 1
                        IncrementAuthorCount(authorCounts, author)
                    End If
                End If

            Case Else
                ' Recurse into child nodes for structure elements like hyperlinks, smart tags, etc.
                For Each child As XmlNode In node.ChildNodes
                    ProcessDocBodyNode(child, nsMgr, sb, includeTrackedChanges,
                                     filterAuthors, filterSince, filterUntil, insCount, delCount, fmtCount, authorCounts)
                Next
        End Select
    End Sub

    Private Shared Function PassesRevisionFilter(author As String, dateStr As String,
                                                  filterAuthors As List(Of String), filterSince As DateTime?, filterUntil As DateTime?) As Boolean
        If filterAuthors IsNot Nothing AndAlso filterAuthors.Count > 0 Then
            Dim authorMatched As Boolean = filterAuthors.Any(
                Function(candidate) Not String.IsNullOrWhiteSpace(candidate) AndAlso
                                    author.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
            If Not authorMatched Then Return False
        End If

        If (filterSince.HasValue OrElse filterUntil.HasValue) AndAlso Not String.IsNullOrWhiteSpace(dateStr) Then
            Dim revDate As DateTime
            If DateTime.TryParse(dateStr, Globalization.CultureInfo.InvariantCulture,
                                 Globalization.DateTimeStyles.AssumeUniversal Or Globalization.DateTimeStyles.AdjustToUniversal, revDate) Then
                If filterSince.HasValue AndAlso revDate < filterSince.Value.ToUniversalTime() Then Return False
                If filterUntil.HasValue AndAlso revDate > filterUntil.Value.ToUniversalTime() Then Return False
            End If
        ElseIf (filterSince.HasValue OrElse filterUntil.HasValue) AndAlso String.IsNullOrWhiteSpace(dateStr) Then
            ' A dated filter cannot deterministically include an undated revision.
            Return False
        End If
        Return True
    End Function

    Private Shared Function TryParseRevisionDateFilter(value As String, isUpperBound As Boolean,
                                                        ByRef parsedValue As DateTime?, ByRef errorMessage As String) As Boolean
        parsedValue = Nothing
        errorMessage = Nothing
        If String.IsNullOrWhiteSpace(value) Then Return True

        Dim trimmed As String = value.Trim()
        Dim parsed As DateTime
        If Not DateTime.TryParse(trimmed, Globalization.CultureInfo.InvariantCulture,
                                 Globalization.DateTimeStyles.AssumeLocal, parsed) Then
            errorMessage = $"Invalid tracked-changes date/time '{trimmed}'. Use ISO 8601, for example 2026-08-24 or 2026-08-24T09:30:00+02:00."
            Return False
        End If

        If isUpperBound AndAlso trimmed.Length = 10 AndAlso trimmed(4) = "-"c AndAlso trimmed(7) = "-"c Then
            parsed = parsed.Date.AddDays(1).AddTicks(-1)
        End If
        parsedValue = parsed
        Return True
    End Function

    Private Sub AppendTrackedChangesXmlDetails(tempDir As System.String, sb As System.Text.StringBuilder,
                                                filterAuthors As System.Collections.Generic.List(Of System.String),
                                                filterSince As System.DateTime?, filterUntil As System.DateTime?)
        Dim wordDirectory As System.String = System.IO.Path.Combine(tempDir, "word")
        If Not System.IO.Directory.Exists(wordDirectory) Then Return

        Dim partPaths As New System.Collections.Generic.List(Of System.String)()
        Dim documentPart As System.String = System.IO.Path.Combine(wordDirectory, "document.xml")
        If System.IO.File.Exists(documentPart) Then partPaths.Add(documentPart)
        partPaths.AddRange(System.IO.Directory.GetFiles(wordDirectory, "header*.xml"))
        partPaths.AddRange(System.IO.Directory.GetFiles(wordDirectory, "footer*.xml"))
        For Each optionalPartName As System.String In New System.String() {"footnotes.xml", "endnotes.xml"}
            Dim optionalPartPath As System.String = System.IO.Path.Combine(wordDirectory, optionalPartName)
            If System.IO.File.Exists(optionalPartPath) Then partPaths.Add(optionalPartPath)
        Next

        Dim details As New System.Text.StringBuilder()
        Dim matchingCount As System.Int32 = 0
        Dim revisionIndex As System.Int32 = 0
        Dim revisionXPath As System.String =
            "//w:ins | //w:del | //w:moveFrom | //w:moveTo | //w:rPrChange | //w:pPrChange | //w:tblPrChange | //w:tblGridChange | //w:trPrChange | //w:tcPrChange | //w:sectPrChange | //w:numberingChange"

        For Each partPath As System.String In partPaths.Distinct(System.StringComparer.OrdinalIgnoreCase)
            Try
                Dim partDocument As New System.Xml.XmlDocument()
                partDocument.Load(partPath)
                Dim partNamespaces As New System.Xml.XmlNamespaceManager(partDocument.NameTable)
                partNamespaces.AddNamespace("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
                Dim revisionNodes As System.Xml.XmlNodeList = partDocument.SelectNodes(revisionXPath, partNamespaces)
                If revisionNodes Is Nothing Then Continue For

                For Each revisionNode As System.Xml.XmlNode In revisionNodes
                    Dim element As System.Xml.XmlElement = TryCast(revisionNode, System.Xml.XmlElement)
                    If element Is Nothing Then Continue For
                    Dim author As System.String = element.GetAttribute("w:author")
                    Dim dateStr As System.String = element.GetAttribute("w:date")
                    If Not PassesRevisionFilter(author, dateStr, filterAuthors, filterSince, filterUntil) Then Continue For

                    matchingCount += 1
                    revisionIndex += 1
                    Dim revisionId As System.String = element.GetAttribute("w:id")
                    Dim revisionText As System.String = ExtractRevisionNodeText(revisionNode, partNamespaces)
                    Dim paragraphContext As System.String = ExtractRevisionParagraphContext(revisionNode, partNamespaces)
                    Dim partName As System.String = System.IO.Path.GetFileName(partPath)

                    details.AppendLine($"[Revision #{revisionIndex}] Part: {partName} | Type: {revisionNode.LocalName} | Author: {If(System.String.IsNullOrWhiteSpace(author), "(unknown)", author)} | Date: {If(System.String.IsNullOrWhiteSpace(dateStr), "(unknown)", dateStr)} | Id: {If(System.String.IsNullOrWhiteSpace(revisionId), "(none)", revisionId)}")
                    If Not System.String.IsNullOrWhiteSpace(revisionText) Then details.AppendLine("  Changed text: " & revisionText)
                    If Not System.String.IsNullOrWhiteSpace(paragraphContext) Then details.AppendLine("  Paragraph context: " & paragraphContext)
                Next
            Catch ex As System.Exception
                ' A malformed optional part must not suppress valid revisions from the remaining parts.
                details.AppendLine($"[Revision part warning] {System.IO.Path.GetFileName(partPath)} could not be inspected: {ex.Message}")
            End Try
        Next

        sb.AppendLine("═══ TRACKED CHANGES XML DETAILS ═══")
        If filterAuthors IsNot Nothing AndAlso filterAuthors.Count > 0 Then
            sb.AppendLine("Author filter: " & System.String.Join(", ", filterAuthors))
        End If
        If filterSince.HasValue Then sb.AppendLine("Since: " & filterSince.Value.ToString("o", System.Globalization.CultureInfo.InvariantCulture))
        If filterUntil.HasValue Then sb.AppendLine("Until: " & filterUntil.Value.ToString("o", System.Globalization.CultureInfo.InvariantCulture))
        sb.AppendLine($"Matching XML revision records: {matchingCount}")
        sb.Append(details)
        sb.AppendLine()
    End Sub

    Private Shared Function ExtractRevisionNodeText(revisionNode As XmlNode, nsMgr As XmlNamespaceManager) As String
        Dim text As New StringBuilder()
        For Each textNode As XmlNode In revisionNode.SelectNodes(".//w:t | .//w:delText | .//w:instrText | .//w:delInstrText", nsMgr)
            text.Append(textNode.InnerText)
        Next
        Return NormalizeRevisionOutputText(text.ToString())
    End Function

    Private Shared Function ExtractRevisionParagraphContext(revisionNode As XmlNode, nsMgr As XmlNamespaceManager) As String
        Dim paragraph As XmlNode = revisionNode
        While paragraph IsNot Nothing AndAlso paragraph.LocalName <> "p"
            paragraph = paragraph.ParentNode
        End While
        If paragraph Is Nothing Then Return Nothing

        Dim text As New StringBuilder()
        For Each textNode As XmlNode In paragraph.SelectNodes(".//w:t | .//w:delText", nsMgr)
            text.Append(textNode.InnerText)
        Next
        Return NormalizeRevisionOutputText(text.ToString())
    End Function

    Private Shared Function NormalizeRevisionOutputText(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return Nothing
        Dim normalized As String = value.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
        If normalized.Length > 1000 Then normalized = normalized.Substring(0, 1000) & "..."
        Return normalized
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
