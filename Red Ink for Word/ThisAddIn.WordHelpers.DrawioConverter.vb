' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WordHelpers.DrawioConverter.vb
' Purpose: Converts a .drawio diagram file into a self-contained interactive
'          HTML flowchart navigator.
'
' Architecture / Key Ideas:
'  - XML Parsing: Reads .drawio XML and extracts mxCell elements, separating
'    vertices (nodes) and edges (arrows) by their attributes.
'  - Object Wrappers: Handles <object>/<UserObject> elements that wrap mxCell
'    children — the id/value live on the wrapper, not the inner cell.
'  - Compressed Diagrams: Handles both raw <mxGraphModel> XML and compressed
'    (base64 + deflate) content stored inside <diagram> elements.
'  - Multi-Page Support: Processes all <diagram> pages in a single .drawio file,
'    merging vertices and edges across pages.
'  - Classification: Classifies each vertex as start (no incoming), decision
'    (incoming + outgoing), end (incoming only), or isolated — matching the
'    original Python reference implementation. Filters unlabeled arrows when
'    labeled ones exist for a source.
'  - Smart Buttons: Single-outgoing nodes show "Start" or "Next" instead of
'    repeating the upcoming node text, avoiding duplicate statements.
'  - Edge Cases: Self-loops, dangling edges, cyclic graphs, multiple start
'    nodes, empty node text, duplicate edges.
'  - HTML Generation: Standalone HTML with embedded JSON, Red Ink logo at
'    bottom, user-provided title, optional Home link, consistent square
'    styling with #B91312 accent.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.ComponentModel.DataAnnotations
Imports System.Diagnostics
Imports System.IO
Imports System.IO.Compression
Imports System.Text
Imports System.Xml
Imports DocumentFormat.OpenXml.EMMA
Imports Google.Apis.Util
Imports Microsoft.Office.Interop.Word
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' =========================================================================
    ' Draw.io Data Models
    ' =========================================================================

    Private Class DrawioVertex
        Public Property Id As String
        Public Property Text As String
        Public Property RiType As String              ' "input_string", "set", "condition", etc.
        Public Property RiAttributes As Dictionary(Of String, String)  ' All ri_* attributes
    End Class

    Private Class DrawioEdge
        Public Property Source As String
        Public Property Target As String
        Public Property Label As String
    End Class

    Private Class DrawioOutgoingArrow
        Public Property TargetId As String
        Public Property Label As String
    End Class

    Private Class DrawioIncomingArrow
        Public Property SourceId As String
        Public Property Label As String
    End Class

    Private Class DrawioElement
        Public Property Id As String
        Public Property Text As String
        Public Property ElementType As String  ' "start", "decision", "end", "isolated"
        Public Property RiType As String       ' "input_string", "set", "condition", etc.
        Public Property RiAttributes As Dictionary(Of String, String)
        Public Property Incoming As List(Of DrawioIncomingArrow)
        Public Property Outgoing As List(Of DrawioOutgoingArrow)
    End Class

    ' =========================================================================
    ' Entry Point
    ' =========================================================================

    ''' <summary>
    ''' Entry point for converting a single .drawio file to interactive HTML.
    ''' Prompts for file selection via DragDropForm, asks for a title and an
    ''' optional Home link, then converts.
    ''' </summary>
    Public Sub ConvertDrawioToHtml()
        Dim selectedPath As String = ""

        Globals.ThisAddIn.DragDropFormLabel = "Select a .drawio file to convert"
        Globals.ThisAddIn.DragDropFormFilter = "Draw.io Diagrams|*.drawio|All Files (*.*)|*.*"

        Try
            Using frm As New DragDropForm(DragDropMode.FileOnly)
                Dim __safeDialogOwner103 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner103 IsNot Nothing, frm.ShowDialog(__safeDialogOwner103), frm.ShowDialog()) = System.Windows.Forms.DialogResult.OK Then
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
        If ext <> ".drawio" Then
            ShowCustomMessageBox($"File type '{ext}' is not supported. Please select a .drawio file.")
            Exit Sub
        End If

        ' Ask user for a title
        Dim defaultTitle As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim userTitle As String = ShowCustomInputBox(
            "Enter a title for the flowchart navigator:",
            AN & " Draw.io Converter", True, defaultTitle)
        If String.IsNullOrWhiteSpace(userTitle) Then Exit Sub
        userTitle = userTitle.Trim()

        ' Ask user for an optional Home link
        Dim homeLink As String = ShowCustomInputBox(
            "Enter a Home URL (leave empty to skip):" & vbCrLf &
            "A Home button will appear so users can navigate back to this URL at any time.",
            AN & " Draw.io Converter", True, "")
        ' ShowCustomInputBox returns "" on cancel for SimpleInput mode — treat both
        ' empty and cancelled as "no home link"
        If homeLink IsNot Nothing Then homeLink = homeLink.Trim()
        If String.IsNullOrEmpty(homeLink) Then homeLink = Nothing

        ' Ask user for an optional custom CSS file
        Dim customCss As String = Nothing
        Dim cssPath As String = ""

        Globals.ThisAddIn.DragDropFormLabel = "Select a custom CSS file (or cancel to use default style)"
        Globals.ThisAddIn.DragDropFormFilter = "CSS Stylesheets|*.css|All Files (*.*)|*.*"

        Try
            Using cssFrm As New DragDropForm(DragDropMode.FileOnly)
                Dim __safeDialogOwner152 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                If If(__safeDialogOwner152 IsNot Nothing, cssFrm.ShowDialog(__safeDialogOwner152), cssFrm.ShowDialog()) = System.Windows.Forms.DialogResult.OK Then
                    cssPath = cssFrm.SelectedFilePath
                End If
            End Using
        Finally
            Globals.ThisAddIn.DragDropFormLabel = ""
            Globals.ThisAddIn.DragDropFormFilter = ""
        End Try

        If Not String.IsNullOrWhiteSpace(cssPath) AndAlso File.Exists(cssPath) Then
            Try
                customCss = File.ReadAllText(cssPath, Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox($"Could not read CSS file:{vbCrLf}{vbCrLf}{ex.Message}", AN & " Draw.io Converter")
            End Try
        End If

        ' Output path: same directory and name, .html extension
        Dim dir As String = Path.GetDirectoryName(selectedPath)
        Dim nameWithoutExt As String = Path.GetFileNameWithoutExtension(selectedPath)
        Dim outputPath As String = Path.Combine(dir, $"{nameWithoutExt}.html")

        Try
            Dim success As Boolean = ConvertSingleDrawioFile(selectedPath, outputPath, userTitle, homeLink, customCss)

            If success Then
                ShowCustomMessageBox($"Converted successfully:{vbCrLf}{vbCrLf}{Path.GetFileName(outputPath)}", AN & " Draw.io Converter")
            Else
                ShowCustomMessageBox($"Conversion failed. The file may not contain a valid flowchart.", AN & " Draw.io Converter")
            End If
        Catch ex As Exception
            ShowCustomMessageBox($"Error during conversion:{vbCrLf}{vbCrLf}{ex.Message}", AN & " Draw.io Converter")
        End Try
    End Sub

    ' =========================================================================
    ' Core Conversion
    ' =========================================================================

    Private Function ConvertSingleDrawioFile(inputPath As String, outputPath As String,
                                             title As String, homeLink As String,
                                             customCss As String) As Boolean
        Try
            Dim content As String = File.ReadAllText(inputPath, Encoding.UTF8)

            Dim vertices As Dictionary(Of String, DrawioVertex) = Nothing
            Dim edges As List(Of DrawioEdge) = Nothing

            ParseDrawio(content, vertices, edges)

            If vertices.Count = 0 Then
                Debug.WriteLine($"ConvertSingleDrawioFile: No vertices found in {Path.GetFileName(inputPath)}")
                Return False
            End If

            Dim elements As List(Of DrawioElement) = ClassifyDrawioElements(vertices, edges)

            If elements.Count = 0 Then
                Debug.WriteLine($"ConvertSingleDrawioFile: No elements after classification in {Path.GetFileName(inputPath)}")
                Return False
            End If

            Dim flowchartData As New JObject()
            flowchartData("elements") = SerializeDrawioElements(elements)

            Dim htmlOutput As String = GenerateDrawioHtml(flowchartData, title, homeLink, customCss)
            File.WriteAllText(outputPath, htmlOutput, Encoding.UTF8)

            Return True

        Catch ex As XmlException
            Debug.WriteLine($"ConvertSingleDrawioFile XML error: {ex.Message}")
            Return False
        Catch ex As Exception
            Debug.WriteLine($"ConvertSingleDrawioFile error: {ex.Message}")
            Return False
        End Try
    End Function

    ' =========================================================================
    ' XML Parsing
    ' =========================================================================

    Private Shared Sub ParseDrawio(content As String,
                                   ByRef vertices As Dictionary(Of String, DrawioVertex),
                                   ByRef edges As List(Of DrawioEdge))

        vertices = New Dictionary(Of String, DrawioVertex)(StringComparer.Ordinal)
        edges = New List(Of DrawioEdge)()

        Dim rootDoc As New XmlDocument()
        rootDoc.LoadXml(content)

        Dim xmlDocsToParse As New List(Of XmlDocument)()

        Dim diagramNodes As XmlNodeList = rootDoc.SelectNodes("//diagram")
        If diagramNodes IsNot Nothing AndAlso diagramNodes.Count > 0 Then
            For Each diagramNode As System.Xml.XmlNode In diagramNodes
                Dim inlineModel As System.Xml.XmlNode = diagramNode.SelectSingleNode("mxGraphModel")
                If inlineModel IsNot Nothing Then
                    Dim inlineDoc As New XmlDocument()
                    inlineDoc.LoadXml(inlineModel.OuterXml)
                    xmlDocsToParse.Add(inlineDoc)
                Else
                    Dim compressedContent As String = diagramNode.InnerText?.Trim()
                    If Not String.IsNullOrEmpty(compressedContent) Then
                        Dim decompressed As String = TryDecompressDrawioContent(compressedContent)
                        If Not String.IsNullOrEmpty(decompressed) Then
                            Try
                                Dim decompDoc As New XmlDocument()
                                decompDoc.LoadXml(decompressed)
                                xmlDocsToParse.Add(decompDoc)
                            Catch ex As XmlException
                                Debug.WriteLine($"ParseDrawio: Failed to parse decompressed diagram: {ex.Message}")
                            End Try
                        End If
                    End If
                End If
            Next
        End If

        If xmlDocsToParse.Count = 0 Then
            xmlDocsToParse.Add(rootDoc)
        End If

        For Each xmlDocToParse As XmlDocument In xmlDocsToParse
            ExtractCellsFromXml(xmlDocToParse, vertices, edges)
        Next
    End Sub

    Private Shared Function TryDecompressDrawioContent(compressed As String) As String
        Try
            Dim bytes As Byte() = System.Convert.FromBase64String(compressed)
            Dim decompressed As String
            Using inputStream As New MemoryStream(bytes)
                Using deflateStream As New DeflateStream(inputStream, CompressionMode.Decompress)
                    Using reader As New StreamReader(deflateStream, Encoding.UTF8)
                        decompressed = reader.ReadToEnd()
                    End Using
                End Using
            End Using
            decompressed = System.Net.WebUtility.UrlDecode(decompressed)
            Return decompressed
        Catch ex As Exception
            Debug.WriteLine($"TryDecompressDrawioContent failed: {ex.Message}")
            Return Nothing
        End Try
    End Function

    Private Shared Sub ExtractCellsFromXml(xmlDoc As XmlDocument,
                                           vertices As Dictionary(Of String, DrawioVertex),
                                           edges As List(Of DrawioEdge))

        ' --- 1. Direct mxCell elements (skip those wrapped by object/UserObject) ---
        For Each cell As System.Xml.XmlNode In xmlDoc.SelectNodes("//mxCell")
            Dim parentName As String = cell.ParentNode?.Name
            If parentName = "object" OrElse parentName = "UserObject" Then Continue For

            Dim cellId As String = cell.Attributes("id")?.Value
            If String.IsNullOrEmpty(cellId) Then Continue For
            If cellId = "0" OrElse cellId = "1" Then Continue For

            If cell.Attributes("vertex")?.Value = "1" Then
                Dim riAttrs As Dictionary(Of String, String) = ExtractRiAttributes(cell)
                vertices(cellId) = New DrawioVertex() With {
                    .Id = cellId,
                    .Text = If(cell.Attributes("value")?.Value, ""),
                    .RiType = If(riAttrs.ContainsKey("ri_type"), riAttrs("ri_type"), ""),
                    .RiAttributes = riAttrs
                }
            ElseIf cell.Attributes("edge")?.Value = "1" Then
                Dim source As String = cell.Attributes("source")?.Value
                Dim target As String = cell.Attributes("target")?.Value
                Dim label As String = If(cell.Attributes("value")?.Value, "")
                If Not String.IsNullOrEmpty(source) AndAlso Not String.IsNullOrEmpty(target) Then
                    edges.Add(New DrawioEdge() With {.Source = source, .Target = target, .Label = label})
                End If
            End If
        Next

        ' --- 2. <object> and <UserObject> wrappers (id/label on wrapper) ---
        For Each wrapperTag As String In {"object", "UserObject"}
            For Each wrapper As System.Xml.XmlNode In xmlDoc.SelectNodes("//" & wrapperTag)
                Dim wrapperId As String = wrapper.Attributes("id")?.Value
                If String.IsNullOrEmpty(wrapperId) Then Continue For
                If wrapperId = "0" OrElse wrapperId = "1" Then Continue For

                Dim wrapperLabel As String = If(wrapper.Attributes("label")?.Value, "")
                Dim innerCell As System.Xml.XmlNode = wrapper.SelectSingleNode("mxCell")
                If innerCell Is Nothing Then Continue For

                If innerCell.Attributes("vertex")?.Value = "1" Then
                    ' Merge ri_* from both wrapper and inner cell (wrapper takes precedence)
                    Dim riAttrs As Dictionary(Of String, String) = ExtractRiAttributes(innerCell)
                    For Each mergeAttr As KeyValuePair(Of String, String) In ExtractRiAttributes(wrapper)
                        riAttrs(mergeAttr.Key) = mergeAttr.Value
                    Next
                    vertices(wrapperId) = New DrawioVertex() With {
                        .Id = wrapperId,
                        .Text = wrapperLabel,
                        .RiType = If(riAttrs.ContainsKey("ri_type"), riAttrs("ri_type"), ""),
                        .RiAttributes = riAttrs
                    }
                ElseIf innerCell.Attributes("edge")?.Value = "1" Then
                    Dim source As String = innerCell.Attributes("source")?.Value
                    Dim target As String = innerCell.Attributes("target")?.Value
                    If Not String.IsNullOrEmpty(source) AndAlso Not String.IsNullOrEmpty(target) Then
                        edges.Add(New DrawioEdge() With {.Source = source, .Target = target, .Label = wrapperLabel})
                    End If
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Extracts all attributes prefixed with "ri_" from an XML node and returns
    ''' them as a case-insensitive dictionary.
    ''' </summary>
    Private Shared Function ExtractRiAttributes(node As System.Xml.XmlNode) As Dictionary(Of String, String)
        Dim result As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If node.Attributes Is Nothing Then Return result
        For Each attr As XmlAttribute In node.Attributes
            If attr.Name.StartsWith("ri_", StringComparison.OrdinalIgnoreCase) Then
                result(attr.Name.ToLowerInvariant()) = attr.Value
            End If
        Next
        Return result
    End Function

    ' =========================================================================
    ' Element Classification — faithful to the Python reference implementation
    ' =========================================================================

    Private Shared Function ClassifyDrawioElements(vertices As Dictionary(Of String, DrawioVertex),
                                                   edges As List(Of DrawioEdge)) As List(Of DrawioElement)

        ' Filter out dangling edges
        Dim validEdges As List(Of DrawioEdge) = edges.Where(
            Function(e) vertices.ContainsKey(e.Source) AndAlso vertices.ContainsKey(e.Target)
        ).ToList()

        ' Group outgoing arrows by source
        Dim outgoingBySource As New Dictionary(Of String, List(Of DrawioEdge))(StringComparer.Ordinal)
        For Each vid As String In vertices.Keys
            outgoingBySource(vid) = New List(Of DrawioEdge)()
        Next
        For Each validEdge As DrawioEdge In validEdges
            If outgoingBySource.ContainsKey(validEdge.Source) Then
                outgoingBySource(validEdge.Source).Add(validEdge)
            End If
        Next

        ' Build filtered outgoing/incoming — skip unlabeled if source has labeled ones
        Dim outgoing As New Dictionary(Of String, List(Of DrawioOutgoingArrow))(StringComparer.Ordinal)
        Dim incoming As New Dictionary(Of String, List(Of DrawioIncomingArrow))(StringComparer.Ordinal)

        For Each vid As String In vertices.Keys
            outgoing(vid) = New List(Of DrawioOutgoingArrow)()
            incoming(vid) = New List(Of DrawioIncomingArrow)()
        Next

        For Each kvp As KeyValuePair(Of String, List(Of DrawioEdge)) In outgoingBySource
            Dim source As String = kvp.Key
            Dim arrows As List(Of DrawioEdge) = kvp.Value
            Dim hasLabeled As Boolean = arrows.Any(Function(a) Not String.IsNullOrEmpty(a.Label))

            For Each arrow As DrawioEdge In arrows
                If Not String.IsNullOrEmpty(arrow.Label) OrElse Not hasLabeled Then
                    outgoing(source).Add(New DrawioOutgoingArrow() With {
                        .TargetId = arrow.Target, .Label = arrow.Label
                    })
                    incoming(arrow.Target).Add(New DrawioIncomingArrow() With {
                        .SourceId = source, .Label = arrow.Label
                    })
                End If
            Next
        Next

        ' Classify — matching Python: has_incoming / has_outgoing booleans
        Dim elements As New List(Of DrawioElement)()

        For Each kvp As KeyValuePair(Of String, DrawioVertex) In vertices
            Dim vid As String = kvp.Key
            Dim vertex As DrawioVertex = kvp.Value

            Dim hasIncoming As Boolean = incoming(vid).Count > 0
            Dim hasOutgoing As Boolean = outgoing(vid).Count > 0

            Dim elementType As String
            If hasOutgoing AndAlso Not hasIncoming Then
                elementType = "start"
            ElseIf hasIncoming AndAlso hasOutgoing Then
                elementType = "decision"
            ElseIf hasIncoming AndAlso Not hasOutgoing Then
                elementType = "end"
            Else
                elementType = "isolated"
            End If

            elements.Add(New DrawioElement() With {
                .Id = vid,
                .Text = vertex.Text,
                .ElementType = elementType,
                .RiType = vertex.RiType,
                .RiAttributes = vertex.RiAttributes,
                .Incoming = If(hasIncoming, incoming(vid), Nothing),
                .Outgoing = If(hasOutgoing, outgoing(vid), Nothing)
            })
        Next

        ' Cyclic graph fallback: promote best candidate to start
        Dim hasStart As Boolean = elements.Any(Function(e) e.ElementType = "start")
        If Not hasStart Then
            Dim candidates = elements.Where(Function(e) e.ElementType = "decision").ToList()
            If candidates.Count = 0 Then
                candidates = elements.Where(Function(e) e.ElementType <> "isolated").ToList()
            End If
            If candidates.Count > 0 Then
                Dim best As DrawioElement = candidates.OrderBy(
                    Function(e) If(e.Incoming IsNot Nothing, e.Incoming.Count, 0)
                ).ThenByDescending(
                    Function(e) If(e.Outgoing IsNot Nothing, e.Outgoing.Count, 0)
                ).First()
                best.ElementType = "start"
            End If
        End If

        Return elements
    End Function

    ' =========================================================================
    ' JSON Serialization
    ' =========================================================================

    Private Shared Function SerializeDrawioElements(elements As List(Of DrawioElement)) As JArray
        Dim arr As New JArray()

        For Each el As DrawioElement In elements
            Dim obj As New JObject()
            obj("id") = el.Id
            obj("text") = el.Text
            obj("type") = el.ElementType

            ' Emit ri_type and all ri_* attributes when present
            If Not String.IsNullOrEmpty(el.RiType) Then
                obj("ri_type") = el.RiType
            End If
            If el.RiAttributes IsNot Nothing AndAlso el.RiAttributes.Count > 0 Then
                Dim riObj As New JObject()
                For Each kvp As KeyValuePair(Of String, String) In el.RiAttributes
                    riObj(kvp.Key) = kvp.Value
                Next
                obj("ri") = riObj
            End If

            If el.Incoming IsNot Nothing AndAlso el.Incoming.Count > 0 Then
                Dim incArr As New JArray()
                For Each inc As DrawioIncomingArrow In el.Incoming
                    Dim incObj As New JObject()
                    incObj("source_id") = inc.SourceId
                    incObj("label") = inc.Label
                    incArr.Add(incObj)
                Next
                obj("incoming") = incArr
            End If

            If el.Outgoing IsNot Nothing AndAlso el.Outgoing.Count > 0 Then
                Dim outArr As New JArray()
                For Each outArrow As DrawioOutgoingArrow In el.Outgoing
                    Dim outObj As New JObject()
                    outObj("target_id") = outArrow.TargetId
                    outObj("label") = outArrow.Label
                    outArr.Add(outObj)
                Next
                obj("outgoing") = outArr
            End If

            arr.Add(obj)
        Next

        Return arr
    End Function

    ' =========================================================================
    ' HTML Generation
    ' =========================================================================

    Private Shared Function GenerateDrawioHtml(flowchartData As JObject, title As String,
                                               homeLink As String, customCss As String) As String
        Dim jsonData As String = flowchartData.ToString(Newtonsoft.Json.Formatting.Indented)
        Dim safeTitle As String = System.Net.WebUtility.HtmlEncode(If(title, "Flowchart Navigator"))
        Dim jsSafeTitle As String = If(title, "Flowchart Navigator").
            Replace("\", "\\").Replace("'", "\'").Replace("""", "\""")

        ' Prepare safe home link for JS — either a quoted string or null
        Dim jsHomeLink As String = "null"
        If Not String.IsNullOrEmpty(homeLink) Then
            jsHomeLink = "'" & homeLink.Replace("\", "\\").Replace("'", "\'") & "'"
        End If

        ' Build custom CSS block if provided
        Dim customCssBlock As String = ""
        If Not String.IsNullOrEmpty(customCss) Then
            customCssBlock = vbCrLf & "    <style>" & vbCrLf & "        /* Custom user styles */" & vbCrLf &
                             "        " & customCss & vbCrLf & "    </style>"
        End If

        Dim html As String = DrawioHtmlTemplate
        html = html.Replace("{{FLOWCHART_DATA}}", jsonData)
        html = html.Replace("{{TITLE}}", safeTitle)
        html = html.Replace("{{JS_TITLE}}", jsSafeTitle)
        html = html.Replace("{{HOME_LINK}}", jsHomeLink)
        html = html.Replace("{{CUSTOM_CSS}}", customCssBlock)
        Return html
    End Function

    Private Const DrawioHtmlTemplate As String =
"<!DOCTYPE html>
<html lang=""en"">
<head>
    <meta charset=""UTF-8"">
    <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">
    <title>{{TITLE}}</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            min-height: 100vh;
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            background: #0a0a0a;
            color: #fafafa;
        }
        .container {
            width: 100%; max-width: 480px; padding: 2rem;
            flex: 1; display: flex; flex-direction: column; justify-content: center;
        }
        .header { text-align: center; margin-bottom: 2rem; }
        .header h1 {
            font-size: 1.1rem; font-weight: 400;
            letter-spacing: 0.05em; color: #ccc;
        }
        .node-text {
            font-size: 1.5rem; font-weight: 300; text-align: center;
            padding: 3rem 1rem; line-height: 1.5;
            border-bottom: 1px solid #222; margin-bottom: 2rem;
        }
        #end-badge { display: none; margin-bottom: 2rem; text-align: center; }
        .end-marker {
            display: inline-block;
            font-size: 0.65rem; font-weight: 500;
            letter-spacing: 0.2em; text-transform: uppercase;
            color: #B91312; padding: 0.5rem 1.5rem;
            border: 1px solid #B91312;
        }
        .choice-buttons { display: flex; flex-direction: column; gap: 0.75rem; }
        .choice-btn {
            background: transparent; border: 1px solid #333;
            color: #fafafa; padding: 1rem 1.5rem;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            font-size: 0.9rem; font-weight: 400;
            cursor: pointer; transition: all 0.15s ease; text-align: left;
        }
        .choice-btn:hover {
            background: #fafafa; color: #0a0a0a; border-color: #fafafa;
        }
        .choice-btn.accent {
            border-color: #B91312; color: #B91312; text-align: center;
        }
        .choice-btn.accent:hover {
            background: #B91312; color: #fafafa; border-color: #B91312;
        }
        .nav-buttons { display: flex; gap: 0.5rem; margin-top: 1rem; }
        .nav-btn {
            display: none; flex: 1;
            background: transparent; border: 1px solid #333;
            color: #666; padding: 0.75rem 1.5rem;
            font-family: Verdana, Geneva, Tahoma, sans-serif;
            font-size: 0.8rem; cursor: pointer;
            transition: all 0.15s ease;
        }
        .nav-btn:hover { color: #fafafa; border-color: #666; }
        .footer {
            text-align: center; padding: 1.5rem 1rem 1rem;
            width: 100%; max-width: 480px;
        }
        .footer-logo { margin-bottom: 0.75rem; }
        .footer-logo img { height: 32px; width: auto; opacity: 0.7; }
        .footer-text { font-size: 0.65rem; color: #444; }
        .footer-text a {
            color: #666; text-decoration: none; transition: color 0.15s ease;
        }
        .footer-text a:hover { color: #fafafa; }
        #isolated-section {
            display: none; margin-top: 2rem;
            padding-top: 1rem; border-top: 1px solid #222;
        }
        #isolated-section h2 {
            font-size: 0.7rem; font-weight: 500;
            letter-spacing: 0.15em; text-transform: uppercase;
            color: #666; margin-bottom: 1rem;
        }
        .isolated-node {
            padding: 0.75rem 1rem; border: 1px solid #222;
            margin-bottom: 0.5rem; font-size: 0.85rem; color: #999;
        }
    </style>{{CUSTOM_CSS}}
</head>
<body>
    <div class=""container"">
        <div class=""header""><h1>{{TITLE}}</h1></div>
        <div id=""node-display"" class=""node-text""></div>
        <div id=""end-badge""><span class=""end-marker"">Complete</span></div>
        <div id=""choice-buttons"" class=""choice-buttons""></div>
        <div class=""nav-buttons"">
            <button id=""back-btn"" class=""nav-btn"">Back</button>
            <button id=""restart-btn"" class=""nav-btn"">Start over</button>
            <button id=""home-btn"" class=""nav-btn"">Home</button>
        </div>
        <div id=""isolated-section"">
            <h2>Additional Notes</h2>
            <div id=""isolated-list""></div>
        </div>
    </div>
    <div class=""footer"">
        <div class=""footer-logo"">
            <a href=""https://redink.ai"" target=""_blank"" rel=""noopener"">
                <img src=""https://redink.ai/content/Red_Ink_Logo_Large.png""
                     alt=""Red Ink"" onerror=""this.parentElement.style.display='none'"">
            </a>
        </div>
        <div class=""footer-text"">
            Created with <a href=""https://redink.ai"" target=""_blank"" rel=""noopener"">Red Ink</a>
        </div>
    </div>
    <script>
        document.title = '{{JS_TITLE}}';
        var homeLink = {{HOME_LINK}};
        var flowchartData = {{FLOWCHART_DATA}};
        var nodeMap = {};
        flowchartData.elements.forEach(function(el) { nodeMap[el.id] = el; });

        /* ===== Variable Store ===== */
        var vars = {};
        var varLabels = {};

        function setVar(name, value) { vars[name] = value; }
        function getVar(name) { return vars.hasOwnProperty(name) ? vars[name] : undefined; }

        /* ===== Resolve dynamic var names like ""rating_{i}"" ===== */
        function resolveDynamic(s) {
            if (!s) return s;
            return s.replace(/\{(\w+)\}/g, function(m, v) {
                var val = getVar(v);
                return (val !== undefined) ? String(val) : m;
            });
        }

        /* ===== Date Helpers ===== */
        function toDate(v) {
            if (!v) return null;
            if (v instanceof Date) return v;
            var d = new Date(v);
            return isNaN(d.getTime()) ? null : d;
        }
        function toISO(d) {
            if (!d) return '';
            if (typeof d === 'string') d = toDate(d);
            if (!d) return '';
            return d.getFullYear() + '-' + ('0'+(d.getMonth()+1)).slice(-2) + '-' + ('0'+d.getDate()).slice(-2);
        }
        function addDaysDate(d, n) {
            var r = new Date(d); r.setDate(r.getDate() + n); return r;
        }
        function addMonthsDate(d, n) {
            var r = new Date(d);
            var day = r.getDate();
            r.setMonth(r.getMonth() + n);
            if (r.getDate() !== day) r.setDate(0);
            return r;
        }
        function endOfMonthDate(d) {
            var r = new Date(d.getFullYear(), d.getMonth() + 1, 0);
            return r;
        }
        function startOfMonthDate(d) {
            return new Date(d.getFullYear(), d.getMonth(), 1);
        }
        function nextMonthEnd(d) {
            var eom = endOfMonthDate(d);
            if (d.getTime() <= eom.getTime() && d.getDate() === eom.getDate()) return eom;
            return endOfMonthDate(addMonthsDate(d, 1));
        }
        function diffDays(a, b) {
            return Math.round((toDate(b) - toDate(a)) / 864e5);
        }
        function fullYears(startD, refD) {
            var y = refD.getFullYear() - startD.getFullYear();
            var mStart = startD.getMonth() * 100 + startD.getDate();
            var mRef = refD.getMonth() * 100 + refD.getDate();
            if (mRef < mStart) y--;
            return y;
        }
        function minDate() {
            var args = [].slice.call(arguments).map(function(a) { return toDate(a); }).filter(Boolean);
            if (args.length === 0) return null;
            return args.reduce(function(a, b) { return a < b ? a : b; });
        }
        function maxDate() {
            var args = [].slice.call(arguments).map(function(a) { return toDate(a); }).filter(Boolean);
            if (args.length === 0) return null;
            return args.reduce(function(a, b) { return a > b ? a : b; });
        }
        /* Union overlapping [start,end] intervals (dates as ms), returns sorted merged array */
        function unionIntervals(intervals) {
            if (!intervals || intervals.length === 0) return [];
            var sorted = intervals.slice().sort(function(a, b) { return a[0] - b[0]; });
            var merged = [sorted[0].slice()];
            for (var i = 1; i < sorted.length; i++) {
                var last = merged[merged.length - 1];
                if (sorted[i][0] <= last[1] + 864e5) {
                    last[1] = Math.max(last[1], sorted[i][1]);
                } else {
                    merged.push(sorted[i].slice());
                }
            }
            return merged;
        }
        /* Count days of merged intervals that overlap [rangeStart, rangeEnd] */
        function overlapDays(intervals, rangeStart, rangeEnd) {
            var rs = rangeStart.getTime(), re = rangeEnd.getTime();
            var total = 0;
            intervals.forEach(function(iv) {
                var s = Math.max(iv[0], rs), e = Math.min(iv[1], re);
                if (s <= e) total += Math.round((e - s) / 864e5) + 1;
            });
            return total;
        }

        /* ===== List Store (for dynamic multi-record collection) ===== */
        /* Lists are stored as vars[name] = [{...}, {...}, ...] */
        function getList(name) {
            var v = getVar(name);
            if (Array.isArray(v)) return v;
            setVar(name, []);
            return vars[name];
        }
        function pushToList(name, record) {
            getList(name).push(record);
        }
        function clearList(name) {
            setVar(name, []);
        }

        /* ===== Safe Expression Evaluator ===== */
        function evalExpr(expr) {
            if (!expr) return undefined;
            var resolved = expr.replace(/\{(\w+)\}/g, function(m, v) {
                var val = getVar(v);
                return (typeof val === 'string') ? JSON.stringify(val) : (val !== undefined ? String(val) : 'undefined');
            });
            try {
                var keys = Object.keys(vars);
                var vals = keys.map(function(k) { return vars[k]; });
                keys.push('round','floor','ceil','abs','min','max',
                           'today','daysBetween','yearsBetween','fullYears',
                           'formatDate','formatNumber','len',
                           'addDays','addMonths','endOfMonth','startOfMonth','nextMonthEnd',
                           'toDate','toISO','minDate','maxDate',
                           'getList','pushToList','clearList',
                           'unionIntervals','overlapDays',
                           'Array','JSON','Math','parseInt','parseFloat','isNaN','String','Number','Boolean');
                vals.push(
                    function(x,d){ d=d||0; var f=Math.pow(10,d); return Math.round(x*f)/f; },
                    Math.floor, Math.ceil, Math.abs, Math.min, Math.max,
                    function(){ return toISO(new Date()); },
                    function(a,b){ return diffDays(a,b); },
                    function(a,b){ return Math.round(diffDays(a,b)/(365.25)*10)/10; },
                    function(a,b){ return fullYears(toDate(a), toDate(b)); },
                    function(d,fmt){ var dt=toDate(d); if(!dt)return ''; if(!fmt)fmt='YYYY-MM-DD';
                        return fmt.replace('YYYY',dt.getFullYear())
                                  .replace('MM',('0'+(dt.getMonth()+1)).slice(-2))
                                  .replace('DD',('0'+dt.getDate()).slice(-2)); },
                    function(n,d){ return Number(n).toFixed(d||0); },
                    function(s){ return s ? (Array.isArray(s) ? s.length : String(s).length) : 0; },
                    function(d,n){ return toISO(addDaysDate(toDate(d),n)); },
                    function(d,n){ return toISO(addMonthsDate(toDate(d),n)); },
                    function(d){ return toISO(endOfMonthDate(toDate(d))); },
                    function(d){ return toISO(startOfMonthDate(toDate(d))); },
                    function(d){ return toISO(nextMonthEnd(toDate(d))); },
                    toDate, toISO, function(){ return toISO(minDate.apply(null,arguments)); },
                    function(){ return toISO(maxDate.apply(null,arguments)); },
                    getList, pushToList, clearList,
                    unionIntervals, overlapDays,
                    Array, JSON, Math, parseInt, parseFloat, isNaN, String, Number, Boolean
                );
                var fn = new Function(keys.join(','), 'return (' + resolved + ');');
                return fn.apply(null, vals);
            } catch(e) {
                console.warn('Expression error:', expr, e);
                return undefined;
            }
        }

        /* ===== Template Resolution ===== */
        function resolveTemplate(tmpl) {
            if (!tmpl) return '';
            return tmpl.replace(/\{([^}]+)\}/g, function(m, inner) {
                var result = evalExpr(inner);
                return (result !== undefined && result !== null) ? String(result) : m;
            }).replace(/\\n/g, '\n');
        }

        /* ===== Loop Stack ===== */
        var loopStack = [];

        var startNode = flowchartData.elements.find(function(el) { return el.type === 'start'; });
        if (!startNode) {
            startNode = flowchartData.elements.find(function(el) { return el.type !== 'isolated'; });
        }
        if (!startNode && flowchartData.elements.length > 0) {
            startNode = flowchartData.elements[0];
        }

        var navHistory = [];
        var nodeDisplay = document.getElementById('node-display');
        var endBadge = document.getElementById('end-badge');
        var choiceButtons = document.getElementById('choice-buttons');
        var backBtn = document.getElementById('back-btn');
        var restartBtn = document.getElementById('restart-btn');
        var homeBtn = document.getElementById('home-btn');

        if (homeLink) {
            homeBtn.style.display = 'inline-block';
            homeBtn.addEventListener('click', function() {
                window.location.href = homeLink;
            });
        }

        function sanitizeText(text) {
            if (!text) return '';
            return text
                .replace(/<br\s*\/?>/gi, '\n')
                .replace(/<\/?div>/gi, '\n')
                .replace(/<\/?p>/gi, '\n')
                .replace(/<\/?span[^>]*>/gi, '')
                .replace(/<\/?b>/gi, '').replace(/<\/?i>/gi, '')
                .replace(/<\/?u>/gi, '').replace(/<\/?font[^>]*>/gi, '')
                .replace(/&amp;/gi, '&').replace(/&lt;/gi, '<')
                .replace(/&gt;/gi, '>').replace(/&quot;/gi, '""')
                .replace(/&#39;/gi, ""'"").replace(/&nbsp;/gi, ' ')
                .replace(/<[^>]*>/g, '').replace(/\n{3,}/g, '\n\n').trim();
        }

        /* ===== Input Form Builder ===== */
        function buildInputField(container, fieldDef) {
            var wrapper = document.createElement('div');
            wrapper.className = 'ri-field-wrapper';
            wrapper.style.cssText = 'margin-bottom:1rem;text-align:left;';

            /* Conditional visibility */
            if (fieldDef.show_if) {
                var visible = !!evalExpr(fieldDef.show_if);
                if (!visible) {
                    wrapper.style.display = 'none';
                    wrapper._hidden = true;
                }
            }

            if (fieldDef.label) {
                var lbl = document.createElement('label');
                lbl.textContent = resolveDynamic(fieldDef.label);
                lbl.style.cssText = 'display:block;font-size:0.75rem;color:#888;margin-bottom:0.3rem;' +
                                    'letter-spacing:0.05em;text-transform:uppercase;';
                wrapper.appendChild(lbl);
            }

            /* Hint text below label */
            if (fieldDef.hint) {
                var hint = document.createElement('div');
                hint.textContent = resolveDynamic(fieldDef.hint);
                hint.style.cssText = 'font-size:0.7rem;color:#666;margin-bottom:0.4rem;font-style:italic;';
                wrapper.appendChild(hint);
            }

            var input;
            var type = (fieldDef.type || 'string').toLowerCase();

            if (type === 'boolean') {
                var checkWrap = document.createElement('div');
                checkWrap.style.cssText = 'display:flex;align-items:center;gap:0.5rem;';
                input = document.createElement('input');
                input.type = 'checkbox';
                input.checked = (fieldDef.default === 'true' || getVar(resolveDynamic(fieldDef.var)) === true);
                input.style.cssText = 'width:1.2rem;height:1.2rem;accent-color:#B91312;flex-shrink:0;';
                input._getVal = function() { return input.checked; };
                checkWrap.appendChild(input);
                if (fieldDef.check_label) {
                    var cLbl = document.createElement('span');
                    cLbl.textContent = resolveDynamic(fieldDef.check_label);
                    cLbl.style.cssText = 'font-size:0.85rem;color:#ccc;cursor:pointer;';
                    cLbl.addEventListener('click', function() { input.checked = !input.checked; });
                    checkWrap.appendChild(cLbl);
                }
                wrapper.appendChild(checkWrap);
                input._fieldDef = fieldDef;
                container.appendChild(wrapper);
                return input;
            } else if (type === 'choice' || type === 'select') {
                input = document.createElement('select');
                input.style.cssText = 'width:100%;padding:0.7rem;background:#111;color:#fafafa;' +
                                      'border:1px solid #333;font-size:0.9rem;font-family:inherit;';
                var opts = (fieldDef.options || '').split(';');
                /* Support value mapping: ""Display Text=value"" */
                opts.forEach(function(o) {
                    var opt = document.createElement('option');
                    var parts = o.trim().split('=');
                    if (parts.length >= 2) {
                        opt.textContent = parts[0].trim();
                        opt.value = parts.slice(1).join('=').trim();
                    } else {
                        opt.value = o.trim();
                        opt.textContent = o.trim();
                    }
                    input.appendChild(opt);
                });
                var prev = getVar(resolveDynamic(fieldDef.var));
                if (prev !== undefined) input.value = String(prev);
                else if (fieldDef.default !== undefined) input.value = String(fieldDef.default);
                input._getVal = function() {
                    var v = input.value;
                    /* Auto-convert numeric strings */
                    if (v !== '' && !isNaN(Number(v)) && fieldDef.value_type !== 'string') return Number(v);
                    if (v === 'true') return true;
                    if (v === 'false') return false;
                    return v;
                };
            } else if (type === 'radio') {
                /* Radio button group */
                var radioDiv = document.createElement('div');
                radioDiv.style.cssText = 'display:flex;flex-direction:column;gap:0.5rem;';
                var radioName = 'radio_' + Math.random().toString(36).substr(2, 9);
                var opts = (fieldDef.options || '').split(';');
                var radios = [];
                opts.forEach(function(o) {
                    var parts = o.trim().split('=');
                    var displayText = parts[0].trim();
                    var val = parts.length >= 2 ? parts.slice(1).join('=').trim() : displayText;
                    var rWrap = document.createElement('label');
                    rWrap.style.cssText = 'display:flex;align-items:center;gap:0.5rem;cursor:pointer;' +
                                          'font-size:0.85rem;color:#ccc;padding:0.3rem 0;';
                    var r = document.createElement('input');
                    r.type = 'radio'; r.name = radioName; r.value = val;
                    r.style.cssText = 'accent-color:#B91312;';
                    var prev = getVar(resolveDynamic(fieldDef.var));
                    if (prev !== undefined && String(prev) === val) r.checked = true;
                    rWrap.appendChild(r);
                    var rText = document.createElement('span');
                    rText.textContent = displayText;
                    rWrap.appendChild(rText);
                    radioDiv.appendChild(rWrap);
                    radios.push(r);
                });
                input = radioDiv;
                input._getVal = function() {
                    var checked = radios.find(function(r) { return r.checked; });
                    if (!checked) return undefined;
                    var v = checked.value;
                    if (v !== '' && !isNaN(Number(v)) && fieldDef.value_type !== 'string') return Number(v);
                    if (v === 'true') return true;
                    if (v === 'false') return false;
                    return v;
                };
                input._fieldDef = fieldDef;
                wrapper.appendChild(radioDiv);
                container.appendChild(wrapper);
                return input;
            } else if (type === 'date') {
                input = document.createElement('input');
                input.type = 'date';
                input.style.cssText = 'width:100%;padding:0.7rem;background:#111;color:#fafafa;' +
                                      'border:1px solid #333;font-size:0.9rem;font-family:inherit;';
                if (fieldDef.min) input.min = fieldDef.min;
                if (fieldDef.max) input.max = fieldDef.max;
                var prev = getVar(resolveDynamic(fieldDef.var));
                if (prev) input.value = prev;
                input._getVal = function() { return input.value; };
            } else {
                input = document.createElement('input');
                input.type = (type === 'integer' || type === 'double') ? 'number' : 'text';
                if (type === 'integer') input.step = '1';
                if (type === 'double') input.step = fieldDef.step || 'any';
                if (fieldDef.min !== undefined) input.min = fieldDef.min;
                if (fieldDef.max !== undefined) input.max = fieldDef.max;
                if (fieldDef.placeholder) input.placeholder = resolveDynamic(fieldDef.placeholder);
                input.style.cssText = 'width:100%;padding:0.7rem;background:#111;color:#fafafa;' +
                                      'border:1px solid #333;font-size:0.9rem;font-family:inherit;';
                var prev = getVar(resolveDynamic(fieldDef.var));
                if (prev !== undefined) input.value = prev;
                else if (fieldDef.default !== undefined) input.value = fieldDef.default;
                input._getVal = function() {
                    if (type === 'integer') return parseInt(input.value, 10) || 0;
                    if (type === 'double') return parseFloat(input.value) || 0;
                    return input.value;
                };
            }

            input._fieldDef = fieldDef;
            wrapper.appendChild(input);
            container.appendChild(wrapper);
            return input;
        }

        /* ===== Build a list editor for dynamic record collection ===== */
        function buildListEditor(container, listDef, onDone) {
            var listName = listDef.var || 'items';
            var fields = [];
            try { fields = JSON.parse(listDef.fields || '[]'); } catch(e) { fields = []; }
            var items = getList(listName);
            var addLabel = listDef.add_label || 'Add entry';
            var doneLabel = listDef.done_label || 'Continue';
            var minItems = parseInt(listDef.min_items || '0', 10);
            var maxItems = parseInt(listDef.max_items || '999', 10);

            function renderList() {
                container.innerHTML = '';

                /* Show existing items */
                items.forEach(function(item, idx) {
                    var row = document.createElement('div');
                    row.style.cssText = 'border:1px solid #333;padding:0.6rem 0.8rem;margin-bottom:0.4rem;' +
                                        'display:flex;justify-content:space-between;align-items:center;font-size:0.8rem;';
                    var text = document.createElement('span');
                    text.style.color = '#ccc';
                    /* Build summary from fields */
                    var parts = [];
                    fields.forEach(function(f) {
                        if (item[f.var] !== undefined && item[f.var] !== '') {
                            parts.push((f.label || f.var) + ': ' + item[f.var]);
                        }
                    });
                    text.textContent = (idx + 1) + '. ' + (parts.join(' | ') || '(entry)');
                    row.appendChild(text);

                    var delBtn = document.createElement('button');
                    delBtn.textContent = '×';
                    delBtn.style.cssText = 'background:transparent;border:1px solid #555;color:#888;' +
                                           'width:1.6rem;height:1.6rem;cursor:pointer;font-size:0.9rem;' +
                                           'display:flex;align-items:center;justify-content:center;flex-shrink:0;';
                    delBtn.addEventListener('click', function() {
                        items.splice(idx, 1);
                        setVar(listName, items);
                        renderList();
                    });
                    row.appendChild(delBtn);
                    container.appendChild(row);
                });

                /* Add-entry form */
                if (items.length < maxItems) {
                    var formDiv = document.createElement('div');
                    formDiv.style.cssText = 'border:1px solid #222;padding:0.8rem;margin-top:0.5rem;';
                    var formInputs = [];
                    fields.forEach(function(f) {
                        formInputs.push(buildInputField(formDiv, f));
                    });

                    var errDiv = document.createElement('div');
                    errDiv.style.cssText = 'color:#B91312;font-size:0.8rem;margin-top:0.3rem;display:none;';
                    formDiv.appendChild(errDiv);

                    var addBtn = document.createElement('button');
                    addBtn.className = 'choice-btn';
                    addBtn.textContent = addLabel;
                    addBtn.style.cssText += 'margin-top:0.5rem;text-align:center;';
                    addBtn.addEventListener('click', function() {
                        errDiv.style.display = 'none';
                        var record = {};
                        for (var i = 0; i < formInputs.length; i++) {
                            var inp = formInputs[i];
                            if (inp._fieldDef && inp._fieldDef.show_if) {
                                var parentWrapper = inp.closest ? inp.closest('.ri-field-wrapper') : inp.parentElement;
                                if (parentWrapper && parentWrapper._hidden) continue;
                            }
                            var fd = inp._fieldDef;
                            if (!fd) continue;
                            var val = inp._getVal();
                            if ((fd.required === 'true' || fd.required === true) &&
                                (val === '' || val === undefined || val === null)) {
                                errDiv.textContent = (fd.label || fd.var) + ' is required.';
                                errDiv.style.display = 'block';
                                return;
                            }
                            record[fd.var] = val;
                        }
                        items.push(record);
                        setVar(listName, items);
                        renderList();
                    });
                    formDiv.appendChild(addBtn);
                    container.appendChild(formDiv);
                }

                /* Done button */
                if (items.length >= minItems) {
                    var doneBtn = document.createElement('button');
                    doneBtn.className = 'choice-btn accent';
                    doneBtn.textContent = doneLabel + (items.length > 0 ? ' (' + items.length + ')' : '');
                    doneBtn.style.cssText += 'margin-top:0.75rem;';
                    doneBtn.addEventListener('click', function() { onDone(); });
                    container.appendChild(doneBtn);
                }
            }
            renderList();
        }

        /* ===== Main Render ===== */
        function renderNode(nodeId, addToHistory) {
            if (addToHistory === undefined) addToHistory = true;
            var node = nodeMap[nodeId];
            if (!node) { nodeDisplay.textContent = 'Error: Node not found'; return; }

            var ri = node.ri || {};
            var riType = (node.ri_type || ri.ri_type || '').toLowerCase();

            /* --- Silent nodes: execute and auto-advance --- */

            if (riType === 'set') {
                var varName = resolveDynamic(ri.ri_var);
                var result = evalExpr(ri.ri_expr);
                if (varName) setVar(varName, result);
                if (node.outgoing && node.outgoing.length > 0) {
                    renderNode(node.outgoing[0].target_id, addToHistory);
                }
                return;
            }

            if (riType === 'set_multi') {
                /* Execute multiple assignments: ri_assignments = ""var1=expr1;var2=expr2;..."" */
                var assignments = (ri.ri_assignments || '').split(';');
                assignments.forEach(function(a) {
                    var eqIdx = a.indexOf('=');
                    if (eqIdx > 0) {
                        var vn = a.substring(0, eqIdx).trim();
                        var ex = a.substring(eqIdx + 1).trim();
                        var res = evalExpr(ex);
                        if (vn) setVar(resolveDynamic(vn), res);
                    }
                });
                if (node.outgoing && node.outgoing.length > 0) {
                    renderNode(node.outgoing[0].target_id, addToHistory);
                }
                return;
            }

            if (riType === 'score') {
                var varName = ri.ri_var || 'score';
                if (getVar(varName) === undefined) setVar(varName, 0);
                var result = evalExpr(ri.ri_expr);
                if (result !== undefined) setVar(varName, result);
                if (node.outgoing && node.outgoing.length > 0) {
                    renderNode(node.outgoing[0].target_id, addToHistory);
                }
                return;
            }

            if (riType === 'condition') {
                var condResult = !!evalExpr(ri.ri_expr);
                if (node.outgoing) {
                    var match = node.outgoing.find(function(a) {
                        var lbl = (a.label || '').toLowerCase().trim();
                        return condResult
                            ? (lbl === 'true' || lbl === 'yes' || lbl === 'ja')
                            : (lbl === 'false' || lbl === 'no' || lbl === 'nein');
                    });
                    if (!match) match = node.outgoing[0];
                    renderNode(match.target_id, addToHistory);
                }
                return;
            }

            if (riType === 'switch') {
                /* Multi-way branch: evaluates ri_expr, follows arrow whose label matches result */
                var switchVal = String(evalExpr(ri.ri_expr) || '');
                if (node.outgoing) {
                    var match = node.outgoing.find(function(a) {
                        return (a.label || '').trim().toLowerCase() === switchVal.toLowerCase();
                    });
                    /* Fallback: arrow labelled ""default"" or ""else"" */
                    if (!match) {
                        match = node.outgoing.find(function(a) {
                            var lbl = (a.label || '').toLowerCase().trim();
                            return lbl === 'default' || lbl === 'else';
                        });
                    }
                    if (!match) match = node.outgoing[0];
                    renderNode(match.target_id, addToHistory);
                }
                return;
            }

            if (riType === 'goto') {
                /* Jump to a node whose id matches the evaluated expression */
                var targetId = String(evalExpr(ri.ri_target) || ri.ri_target || '');
                if (nodeMap[targetId]) {
                    renderNode(targetId, addToHistory);
                } else if (node.outgoing && node.outgoing.length > 0) {
                    renderNode(node.outgoing[0].target_id, addToHistory);
                }
                return;
            }

            if (riType === 'loop') {
                var loopVar = ri.ri_var || '_i';
                var from = parseInt(evalExpr(ri.ri_from || '1'), 10) || 1;
                var to = parseInt(evalExpr(ri.ri_to || '1'), 10) || 1;
                var step = parseInt(ri.ri_step || '1', 10) || 1;
                var current = getVar(loopVar);
                if (current === undefined) {
                    setVar(loopVar, from);
                    current = from;
                } else {
                    current += step;
                    setVar(loopVar, current);
                }
                var bodyArrow = node.outgoing && node.outgoing.find(function(a) {
                    var lbl = (a.label || '').toLowerCase().trim();
                    return lbl === 'body' || lbl === 'loop';
                });
                var doneArrow = node.outgoing && node.outgoing.find(function(a) {
                    var lbl = (a.label || '').toLowerCase().trim();
                    return lbl === 'done' || lbl === 'exit';
                });
                if (current <= to && bodyArrow) {
                    loopStack.push({ nodeId: nodeId, varName: loopVar });
                    renderNode(bodyArrow.target_id, addToHistory);
                } else {
                    setVar(loopVar, undefined);
                    if (doneArrow) renderNode(doneArrow.target_id, addToHistory);
                }
                return;
            }

            if (riType === 'foreach') {
                /* Iterate over a list variable. ri_var = counter, ri_list = list var name */
                var listName = ri.ri_list;
                var iterVar = ri.ri_var || '_item';
                var idxVar = ri.ri_index || '_idx';
                var list = getList(listName);
                var idx = getVar(idxVar);
                if (idx === undefined) {
                    idx = 0;
                } else {
                    idx++;
                }
                setVar(idxVar, idx);
                if (idx < list.length) {
                    setVar(iterVar, list[idx]);
                    /* Also expand item properties as top-level vars for easy access */
                    var item = list[idx];
                    if (item && typeof item === 'object') {
                        Object.keys(item).forEach(function(k) {
                            setVar('_' + k, item[k]);
                        });
                    }
                    var bodyArrow = node.outgoing && node.outgoing.find(function(a) {
                        var lbl = (a.label || '').toLowerCase().trim();
                        return lbl === 'body' || lbl === 'loop';
                    });
                    if (bodyArrow) {
                        loopStack.push({ nodeId: nodeId, varName: idxVar, foreachVar: iterVar });
                        renderNode(bodyArrow.target_id, addToHistory);
                    }
                } else {
                    setVar(idxVar, undefined);
                    setVar(iterVar, undefined);
                    var doneArrow = node.outgoing && node.outgoing.find(function(a) {
                        var lbl = (a.label || '').toLowerCase().trim();
                        return lbl === 'done' || lbl === 'exit';
                    });
                    if (doneArrow) renderNode(doneArrow.target_id, addToHistory);
                }
                return;
            }

            if (riType === 'loop_exit') {
                if (loopStack.length > 0) {
                    var loop = loopStack.pop();
                    setVar(loop.varName, undefined);
                    if (loop.foreachVar) setVar(loop.foreachVar, undefined);
                    var loopNode = nodeMap[loop.nodeId];
                    var doneArrow = loopNode && loopNode.outgoing && loopNode.outgoing.find(function(a) {
                        var lbl = (a.label || '').toLowerCase().trim();
                        return lbl === 'done' || lbl === 'exit';
                    });
                    if (doneArrow) renderNode(doneArrow.target_id, addToHistory);
                }
                return;
            }

            /* --- Visible nodes: render UI --- */

            if (addToHistory && navHistory[navHistory.length - 1] !== nodeId) {
                navHistory.push(nodeId);
            }

            var displayText = sanitizeText(node.text);
            choiceButtons.innerHTML = '';
            backBtn.style.display = navHistory.length > 1 ? 'inline-block' : 'none';
            endBadge.style.display = 'none';
            restartBtn.style.display = 'none';
            nodeDisplay.style.whiteSpace = '';

            /* --- Section header node --- */
            if (riType === 'section') {
                nodeDisplay.style.cssText += 'font-size:1.1rem;font-weight:500;letter-spacing:0.05em;' +
                    'color:#B91312;border-bottom:2px solid #B91312;padding-bottom:1rem;text-transform:uppercase;';
                nodeDisplay.textContent = displayText || '...';
                if (node.outgoing && node.outgoing.length > 0) {
                    var nextBtn = document.createElement('button');
                    nextBtn.className = 'choice-btn accent';
                    nextBtn.textContent = 'Continue';
                    nextBtn.addEventListener('click', function() {
                        renderNode(node.outgoing[0].target_id);
                    });
                    choiceButtons.appendChild(nextBtn);
                }
                return;
            }

            /* --- Input list (dynamic multi-record collection) --- */
            if (riType === 'input_list') {
                nodeDisplay.textContent = displayText || ri.ri_label || '...';
                if (ri.ri_hint) {
                    var hintDiv = document.createElement('div');
                    hintDiv.textContent = resolveDynamic(ri.ri_hint);
                    hintDiv.style.cssText = 'font-size:0.75rem;color:#666;font-style:italic;margin-top:0.5rem;';
                    nodeDisplay.appendChild(hintDiv);
                }
                buildListEditor(choiceButtons, {
                    var: ri.ri_var || 'items',
                    fields: ri.ri_fields,
                    add_label: ri.ri_add_label,
                    done_label: ri.ri_done_label,
                    min_items: ri.ri_min_items,
                    max_items: ri.ri_max_items
                }, function() {
                    if (node.outgoing && node.outgoing.length > 0) {
                        renderNode(node.outgoing[0].target_id);
                    }
                });
                return;
            }

            /* --- Input nodes --- */
            if (riType.indexOf('input') === 0) {
                nodeDisplay.textContent = displayText || ri.ri_label || '...';
                if (ri.ri_hint) {
                    nodeDisplay.style.whiteSpace = 'pre-wrap';
                    nodeDisplay.textContent += '\n';
                    var hintSpan = document.createElement('div');
                    hintSpan.textContent = resolveDynamic(ri.ri_hint);
                    hintSpan.style.cssText = 'font-size:0.75rem;color:#666;font-style:italic;margin-top:0.5rem;';
                    nodeDisplay.appendChild(hintSpan);
                }

                var fields = [];
                if (riType === 'input_group') {
                    try { fields = JSON.parse(ri.ri_fields || '[]'); } catch(e) { fields = []; }
                } else {
                    fields = [{
                        var: resolveDynamic(ri.ri_var) || 'input',
                        type: riType.replace('input_', ''),
                        label: ri.ri_label || '',
                        placeholder: ri.ri_placeholder || '',
                        required: ri.ri_required,
                        min: ri.ri_min, max: ri.ri_max,
                        step: ri.ri_step,
                        default: ri.ri_default,
                        options: ri.ri_options,
                        pattern: ri.ri_pattern,
                        hint: ri.ri_field_hint,
                        show_if: ri.ri_show_if,
                        check_label: ri.ri_check_label,
                        value_type: ri.ri_value_type
                    }];
                }

                var inputEls = [];
                var formDiv = document.createElement('div');
                formDiv.style.cssText = 'margin-bottom:1rem;';

                /* Re-evaluate conditional visibility when fields change */
                var allInputs = [];
                fields.forEach(function(f) {
                    var inp = buildInputField(formDiv, f);
                    inputEls.push(inp);
                    allInputs.push(inp);
                });

                /* Wire up show_if re-evaluation on change events */
                allInputs.forEach(function(inp) {
                    var handler = function() {
                        /* Temporarily store current values */
                        allInputs.forEach(function(ai) {
                            if (ai._fieldDef && ai._getVal) {
                                var tmpName = '__tmp_' + (ai._fieldDef.var || '');
                                vars[tmpName] = ai._getVal();
                            }
                        });
                        /* Re-check visibility */
                        allInputs.forEach(function(ai) {
                            if (ai._fieldDef && ai._fieldDef.show_if) {
                                /* Map __tmp_ vars to real names for eval */
                                allInputs.forEach(function(x) {
                                    if (x._fieldDef) vars[x._fieldDef.var] = vars['__tmp_' + x._fieldDef.var];
                                });
                                var vis = !!evalExpr(ai._fieldDef.show_if);
                                var w = ai.closest ? ai.closest('.ri-field-wrapper') : ai.parentElement;
                                if (w) {
                                    w.style.display = vis ? '' : 'none';
                                    w._hidden = !vis;
                                }
                                /* Cleanup */
                                allInputs.forEach(function(x) {
                                    if (x._fieldDef) delete vars['__tmp_' + x._fieldDef.var];
                                    if (x._fieldDef) delete vars[x._fieldDef.var];
                                });
                            }
                        });
                    };
                    if (inp.addEventListener) inp.addEventListener('change', handler);
                });

                choiceButtons.appendChild(formDiv);

                var errDiv = document.createElement('div');
                errDiv.style.cssText = 'color:#B91312;font-size:0.8rem;margin-bottom:0.5rem;display:none;';
                choiceButtons.appendChild(errDiv);

                var submitBtn = document.createElement('button');
                submitBtn.className = 'choice-btn accent';
                submitBtn.textContent = ri.ri_submit_label || 'Next';
                submitBtn.addEventListener('click', function() {
                    errDiv.style.display = 'none';
                    for (var i = 0; i < inputEls.length; i++) {
                        var inp = inputEls[i];
                        var fd = inp._fieldDef;
                        if (!fd) continue;

                        /* Skip hidden conditional fields */
                        var w = inp.closest ? inp.closest('.ri-field-wrapper') : inp.parentElement;
                        if (w && w._hidden) continue;

                        var val = inp._getVal();

                        if ((fd.required === 'true' || fd.required === true) &&
                            (val === '' || val === undefined || val === null)) {
                            errDiv.textContent = (fd.label || fd.var) + ' is required.';
                            errDiv.style.display = 'block';
                            return;
                        }

                        if (fd.pattern && typeof val === 'string') {
                            var re = new RegExp(fd.pattern);
                            if (!re.test(val)) {
                                errDiv.textContent = (fd.label || fd.var) + ' format is invalid.';
                                errDiv.style.display = 'block';
                                return;
                            }
                        }

                        setVar(resolveDynamic(fd.var), val);
                        if (fd.label) varLabels[resolveDynamic(fd.var)] = resolveDynamic(fd.label);
                    }

                    if (node.outgoing && node.outgoing.length > 0) {
                        renderNode(node.outgoing[0].target_id);
                    }
                });
                choiceButtons.appendChild(submitBtn);
                return;
            }

            /* --- Validate node --- */
            if (riType === 'validate') {
                var valid = !!evalExpr(ri.ri_expr);
                if (valid) {
                    if (node.outgoing) {
                        var passArrow = node.outgoing.find(function(a) {
                            return (a.label||'').toLowerCase().trim() === 'pass';
                        }) || node.outgoing[0];
                        renderNode(passArrow.target_id, addToHistory);
                    }
                } else {
                    nodeDisplay.textContent = resolveTemplate(ri.ri_message || 'Validation failed.');
                    var retryBtn = document.createElement('button');
                    retryBtn.className = 'choice-btn accent';
                    retryBtn.textContent = 'Go back';
                    retryBtn.addEventListener('click', function() {
                        if (navHistory.length > 1) {
                            navHistory.pop();
                            renderNode(navHistory[navHistory.length - 1], false);
                        }
                    });
                    choiceButtons.appendChild(retryBtn);
                }
                return;
            }

            /* --- Output node --- */
            if (riType === 'output') {
                var resolved = resolveTemplate(ri.ri_template || displayText);
                nodeDisplay.style.whiteSpace = 'pre-wrap';
                nodeDisplay.textContent = resolved;
                if (node.outgoing && node.outgoing.length > 0) {
                    var nextBtn = document.createElement('button');
                    nextBtn.className = 'choice-btn accent';
                    nextBtn.textContent = 'Next';
                    nextBtn.addEventListener('click', function() {
                        renderNode(node.outgoing[0].target_id);
                    });
                    choiceButtons.appendChild(nextBtn);
                } else {
                    endBadge.style.display = 'block';
                    restartBtn.style.display = 'inline-block';
                }
                return;
            }

            /* --- Summary node --- */
            if (riType === 'summary') {
                if (ri.ri_template) {
                    nodeDisplay.style.whiteSpace = 'pre-wrap';
                    nodeDisplay.textContent = resolveTemplate(ri.ri_template);
                } else {
                    nodeDisplay.innerHTML = '';
                    var table = document.createElement('table');
                    table.style.cssText = 'width:100%;border-collapse:collapse;text-align:left;';
                    Object.keys(varLabels).forEach(function(vn) {
                        var tr = document.createElement('tr');
                        var tdL = document.createElement('td');
                        tdL.style.cssText = 'padding:0.4rem 0.5rem;color:#888;font-size:0.8rem;border-bottom:1px solid #222;';
                        tdL.textContent = varLabels[vn];
                        var tdV = document.createElement('td');
                        tdV.style.cssText = 'padding:0.4rem 0.5rem;color:#fafafa;font-size:0.9rem;border-bottom:1px solid #222;';
                        var v = getVar(vn);
                        tdV.textContent = (v !== undefined && v !== null) ? String(v) : '';
                        tr.appendChild(tdL); tr.appendChild(tdV);
                        table.appendChild(tr);
                    });
                    nodeDisplay.appendChild(table);
                }
                if (node.outgoing && node.outgoing.length > 0) {
                    var nextBtn = document.createElement('button');
                    nextBtn.className = 'choice-btn accent';
                    nextBtn.textContent = 'Next';
                    nextBtn.addEventListener('click', function() {
                        renderNode(node.outgoing[0].target_id);
                    });
                    choiceButtons.appendChild(nextBtn);
                } else {
                    endBadge.style.display = 'block';
                    restartBtn.style.display = 'inline-block';
                }
                return;
            }

            /* --- Default: original navigation behavior --- */
            nodeDisplay.textContent = displayText || '...';
            var isEnd = (node.type === 'end' || !node.outgoing || node.outgoing.length === 0);

            if (isEnd) {
                endBadge.style.display = 'block';
                restartBtn.style.display = 'inline-block';
            } else {
                var seen = {};
                var uniqueOutgoing = node.outgoing.filter(function(arrow) {
                    var key = arrow.target_id + '|' + (arrow.label || '');
                    if (seen[key]) return false;
                    seen[key] = true;
                    return true;
                });
                uniqueOutgoing.sort(function(a, b) {
                    var labelA = (a.label || '').toLowerCase();
                    var labelB = (b.label || '').toLowerCase();
                    return labelA < labelB ? 1 : labelA > labelB ? -1 : 0;
                });
                var isSinglePath = (uniqueOutgoing.length === 1);
                var isFirstNode = (navHistory.length <= 1);
                uniqueOutgoing.forEach(function(arrow) {
                    var btn = document.createElement('button');
                    btn.className = 'choice-btn';
                    if (arrow.label) {
                        btn.textContent = arrow.label;
                    } else if (isSinglePath) {
                        btn.textContent = isFirstNode ? 'Start' : 'Next';
                        btn.className = 'choice-btn accent';
                    } else {
                        var targetNode = nodeMap[arrow.target_id];
                        btn.textContent = (targetNode && sanitizeText(targetNode.text))
                            ? sanitizeText(targetNode.text) : 'Next';
                    }
                    btn.addEventListener('click', function() { renderNode(arrow.target_id); });
                    choiceButtons.appendChild(btn);
                });
            }
        }

        backBtn.addEventListener('click', function() {
            if (navHistory.length > 1) {
                navHistory.pop();
                renderNode(navHistory[navHistory.length - 1], false);
            }
        });
        restartBtn.addEventListener('click', function() {
            navHistory.length = 0;
            vars = {};
            varLabels = {};
            loopStack = [];
            if (startNode) renderNode(startNode.id);
        });

        var isolatedNodes = flowchartData.elements.filter(function(el) {
            return el.type === 'isolated';
        });
        if (isolatedNodes.length > 0) {
            var isolatedSection = document.getElementById('isolated-section');
            var isolatedList = document.getElementById('isolated-list');
            isolatedNodes.forEach(function(el) {
                var nodeText = sanitizeText(el.text);
                if (nodeText) {
                    var div = document.createElement('div');
                    div.className = 'isolated-node';
                    div.textContent = nodeText;
                    isolatedList.appendChild(div);
                }
            });
            if (isolatedList.children.length > 0) {
                isolatedSection.style.display = 'block';
            }
        }

        if (startNode) { renderNode(startNode.id); }
        else { nodeDisplay.textContent = 'No flowchart nodes found.'; }
    </script>
</body>
</html>"

End Class
