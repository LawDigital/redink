' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Slides.vb
' Purpose: Provides PowerPoint presentation manipulation capabilities using
'          OpenXML SDK for slide extraction, creation, and modification.
'
' Architecture:
'  - JSON Extraction: Reads .pptx files and serializes slide metadata, layouts,
'    placeholders, and text content to JSON for AI processing.
'  - Plan Application: Parses AI-generated JSON action plans and executes
'    slide additions with titles, text, bullets, shapes, and SVG icons.
'  - Layout Resolution: Resolves layouts by relId, URI, or name; falls back
'    to cover-like or default layouts when not found.
'  - Template Cloning: Clones layout structure to new slides, copies images,
'    and purges sample text from placeholders.
'  - Text Population: Supports plain text, bulleted lists, and freestanding
'    text boxes with style properties (font, size, bold, italic, color).
'  - Shape Creation: Adds geometric shapes with fill, outline, and optional
'    text content using absolute or percentage-based positioning.
'  - Speaker Notes: Creates or updates notes slide parts with provided content.
'  - Validation: Includes PPTX package validation and OpenXML error checking.
'  - External Dependencies: DocumentFormat.OpenXml, System.Text.Json,
'    SharedLibrary.SharedMethods for UI messaging.
' =============================================================================


Option Explicit On
Option Strict Off

Imports System.Data
Imports System.Diagnostics
Imports System.Text.Json
Imports System.Text.Json.Serialization
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Validation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn


#Region "JSON Extraction"

    ''' <summary>
    ''' Extracts presentation metadata and content from a PPTX file as JSON.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file.</param>
    ''' <returns>JSON string containing slides, layouts, and content; empty string on error.</returns>
    Public Function GetPresentationJson(pptxPath As String) As String
        If Not System.IO.File.Exists(pptxPath) Then
            ShowCustomMessageBox($"File not found: {pptxPath}")
            Return String.Empty
        End If

        If Not IsValidPptxPackage(pptxPath) Then
            ShowCustomMessageBox("PowerPoint file is corrupt or unreadable.")
            Return String.Empty
        End If

        Try
            Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
                DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)

                Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                If presPart Is Nothing OrElse presPart.Presentation Is Nothing Then
                    ShowCustomMessageBox("Invalid or corrupted presentation.")
                    Return String.Empty
                End If

                Dim result As New PresentationJson With {
                    .Title = presDoc.PackageProperties.Title,
                    .Slides = New List(Of SlideJson)(),
                    .Layouts = New List(Of LayoutJson)()
                }

                If presPart.Presentation.SlideSize IsNot Nothing AndAlso
                   presPart.Presentation.SlideSize.Cx IsNot Nothing AndAlso
                   presPart.Presentation.SlideSize.Cy IsNot Nothing Then
                    result.SlideSize = New SlideSizeJson With {
                        .Width = presPart.Presentation.SlideSize.Cx.Value,
                        .Height = presPart.Presentation.SlideSize.Cy.Value
                    }
                End If

                Dim slideIdList = presPart.Presentation.SlideIdList
                Dim hasSlides As Boolean =
                    (slideIdList IsNot Nothing AndAlso
                     slideIdList.ChildElements.OfType(Of DocumentFormat.OpenXml.Presentation.SlideId)().Any())

                Dim jsonOptions As New System.Text.Json.JsonSerializerOptions With {
                    .WriteIndented = False,
                    .DefaultIgnoreCondition = System.Text.Json.Serialization.JsonIgnoreCondition.WhenWritingNull
                }

                If Not hasSlides Then
                    Try
                        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
                            If sm Is Nothing Then Continue For

                            Dim masterName As System.String = GetMasterName(sm)

                            For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                                If layoutPart Is Nothing OrElse layoutPart.Uri Is Nothing Then Continue For

                                Dim name As System.String = GetLayoutName(layoutPart)
                                Dim layoutUri As System.String = layoutPart.Uri.ToString()
                                Dim relId As System.String = System.String.Empty
                                Dim placeholderDetails As New List(Of PlaceholderJson)

                                Try
                                    relId = sm.GetIdOfPart(layoutPart)
                                Catch
                                End Try

                                If layoutPart.SlideLayout IsNot Nothing AndAlso
                                   layoutPart.SlideLayout.CommonSlideData IsNot Nothing AndAlso
                                   layoutPart.SlideLayout.CommonSlideData.ShapeTree IsNot Nothing Then

                                    CollectDetailedPlaceholdersFromShapeTree(
                                        layoutPart.SlideLayout.CommonSlideData.ShapeTree,
                                        placeholderDetails,
                                        includeText:=False)

                                    MarkPrimaryBodyPlaceholder(
                                        placeholderDetails,
                                        preferText:=False)
                                End If

                                result.Layouts.Add(New LayoutJson With {
                                    .Name = name,
                                    .LayoutId = layoutUri,
                                    .LayoutRelId = relId,
                                    .Master = masterName,
                                    .PlaceholderDetails = placeholderDetails
                                })
                            Next
                        Next
                    Catch
                    End Try

                    EnrichPresentationJson(presPart, result)
                    Return System.Text.Json.JsonSerializer.Serialize(result, jsonOptions)
                End If

                Try
                    Dim idx As Integer = 0

                    For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In
                        slideIdList.ChildElements.OfType(Of DocumentFormat.OpenXml.Presentation.SlideId)()

                        If sid.RelationshipId Is Nothing Then Continue For

                        Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = Nothing
                        Try
                            sp = TryCast(
                                presPart.GetPartById(sid.RelationshipId),
                                DocumentFormat.OpenXml.Packaging.SlidePart)
                        Catch
                            Continue For
                        End Try

                        If sp Is Nothing Then Continue For

                        Dim title As String = GetSlideTitle(sp)
                        Dim key As String = If(
                            String.IsNullOrWhiteSpace(title),
                            $"SID-{sid.Id.Value}",
                            $"{SanitizeKey(title)}-{sid.Id.Value}"
                        )

                        Dim layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp.SlideLayoutPart
                        Dim layoutName As String = GetLayoutName(layoutPart)
                        Dim masterName As String = If(
                            layoutPart IsNot Nothing,
                            GetMasterName(layoutPart.SlideMasterPart),
                            String.Empty
                        )

                        Dim placeholders As New List(Of String)
                        Dim content As New List(Of String)
                        Dim placeholderDetails As New List(Of PlaceholderJson)

                        If sp.Slide IsNot Nothing AndAlso
                           sp.Slide.CommonSlideData IsNot Nothing AndAlso
                           sp.Slide.CommonSlideData.ShapeTree IsNot Nothing Then

                            CollectPlaceholdersFromShapeTree(
                                sp.Slide.CommonSlideData.ShapeTree,
                                placeholders)

                            CollectTextsFromShapeTree(
                                sp.Slide.CommonSlideData.ShapeTree,
                                content)

                            CollectDetailedPlaceholdersFromShapeTree(
                                sp.Slide.CommonSlideData.ShapeTree,
                                placeholderDetails,
                                includeText:=True)

                            MarkPrimaryBodyPlaceholder(
                                placeholderDetails,
                                preferText:=True)
                        End If

                        result.Slides.Add(New SlideJson With {
                            .SlideKey = key,
                            .SlideId = sid.Id.Value,
                            .Index = idx,
                            .Title = title,
                            .Layout = layoutName,
                            .Master = masterName,
                            .Placeholders = placeholders,
                            .Content = content,
                            .PlaceholderDetails = placeholderDetails
                        })

                        idx += 1
                    Next
                Catch
                    EnrichPresentationJson(presPart, result)
                    Return System.Text.Json.JsonSerializer.Serialize(result, jsonOptions)
                End Try

                Try
                    For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
                        If sm Is Nothing Then Continue For

                        Dim masterName As String = GetMasterName(sm)

                        For Each layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                            If layoutPart Is Nothing OrElse layoutPart.Uri Is Nothing Then Continue For

                            Dim name As String = GetLayoutName(layoutPart)
                            Dim layoutUri As String = layoutPart.Uri.ToString()
                            Dim relId As String = String.Empty
                            Dim placeholderDetails As New List(Of PlaceholderJson)

                            Try
                                relId = sm.GetIdOfPart(layoutPart)
                            Catch
                            End Try

                            If layoutPart.SlideLayout IsNot Nothing AndAlso
                               layoutPart.SlideLayout.CommonSlideData IsNot Nothing AndAlso
                               layoutPart.SlideLayout.CommonSlideData.ShapeTree IsNot Nothing Then

                                CollectDetailedPlaceholdersFromShapeTree(
                                    layoutPart.SlideLayout.CommonSlideData.ShapeTree,
                                    placeholderDetails,
                                    includeText:=False)

                                MarkPrimaryBodyPlaceholder(
                                    placeholderDetails,
                                    preferText:=False)
                            End If

                            result.Layouts.Add(New LayoutJson With {
                                .Name = name,
                                .LayoutId = layoutUri,
                                .LayoutRelId = relId,
                                .Master = masterName,
                                .PlaceholderDetails = placeholderDetails
                            })
                        Next
                    Next
                Catch
                End Try

                EnrichPresentationJson(presPart, result)
                Return System.Text.Json.JsonSerializer.Serialize(result, jsonOptions)
            End Using

        Catch ex As System.IO.IOException
            ShowCustomMessageBox($"Error opening presentation (I/O): {ex.Message}")
            Return String.Empty
        Catch ex As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ShowCustomMessageBox($"Error processing presentation (OpenXML): {ex.Message}")
            Return String.Empty
        Catch ex As System.Exception
            ShowCustomMessageBox($"Unexpected error: {ex.Message}")
            Return String.Empty
        End Try
    End Function


    Private Shared Sub CollectDetailedPlaceholdersFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal placeholders As System.Collections.Generic.List(Of PlaceholderJson),
        Optional ByVal includeText As System.Boolean = True)

        If tree Is Nothing OrElse placeholders Is Nothing Then Return

        Dim sourceOrder As Integer = 0

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements
            CollectDetailedPlaceholdersFromElement(child, placeholders, includeText, sourceOrder)
        Next
    End Sub

    Private Shared Sub CollectDetailedPlaceholdersFromElement(
        ByVal child As DocumentFormat.OpenXml.OpenXmlElement,
        ByVal placeholders As System.Collections.Generic.List(Of PlaceholderJson),
        ByVal includeText As System.Boolean,
        ByRef sourceOrder As Integer)

        If child Is Nothing Then Return

        If TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
            Dim grp As DocumentFormat.OpenXml.Presentation.GroupShape =
                CType(child, DocumentFormat.OpenXml.Presentation.GroupShape)

            For Each inner As DocumentFormat.OpenXml.OpenXmlElement In grp.ChildElements
                CollectDetailedPlaceholdersFromElement(inner, placeholders, includeText, sourceOrder)
            Next

            Return
        End If

        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = Nothing
        Dim name As String = String.Empty
        Dim shapeId As Nullable(Of UInteger) = Nothing
        Dim kind As String = String.Empty
        Dim x As Nullable(Of Long) = Nothing
        Dim y As Nullable(Of Long) = Nothing
        Dim cx As Nullable(Of Long) = Nothing
        Dim cy As Nullable(Of Long) = Nothing

        If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
            Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                CType(child, DocumentFormat.OpenXml.Presentation.Shape)

            ph = shp.NonVisualShapeProperties?.
                ApplicationNonVisualDrawingProperties?.
                PlaceholderShape

            name = If(
                shp.NonVisualShapeProperties?.
                    NonVisualDrawingProperties?.
                    Name?.
                    Value,
                String.Empty)

            If shp.NonVisualShapeProperties IsNot Nothing AndAlso
               shp.NonVisualShapeProperties.NonVisualDrawingProperties IsNot Nothing AndAlso
               shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id IsNot Nothing Then
                shapeId = shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value
            End If

            kind = "shape"

            Dim xfrm = shp.ShapeProperties?.Transform2D
            If xfrm IsNot Nothing Then
                If xfrm.Offset IsNot Nothing Then
                    If xfrm.Offset.X IsNot Nothing Then x = xfrm.Offset.X.Value
                    If xfrm.Offset.Y IsNot Nothing Then y = xfrm.Offset.Y.Value
                End If
                If xfrm.Extents IsNot Nothing Then
                    If xfrm.Extents.Cx IsNot Nothing Then cx = xfrm.Extents.Cx.Value
                    If xfrm.Extents.Cy IsNot Nothing Then cy = xfrm.Extents.Cy.Value
                End If
            End If

        ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.Picture Then
            Dim pic As DocumentFormat.OpenXml.Presentation.Picture =
                CType(child, DocumentFormat.OpenXml.Presentation.Picture)

            ph = pic.NonVisualPictureProperties?.
                ApplicationNonVisualDrawingProperties?.
                PlaceholderShape

            name = If(
                pic.NonVisualPictureProperties?.
                    NonVisualDrawingProperties?.
                    Name?.
                    Value,
                String.Empty)

            If pic.NonVisualPictureProperties IsNot Nothing AndAlso
               pic.NonVisualPictureProperties.NonVisualDrawingProperties IsNot Nothing AndAlso
               pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id IsNot Nothing Then
                shapeId = pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id.Value
            End If

            kind = "picture"

            Dim xfrm = pic.ShapeProperties?.Transform2D
            If xfrm IsNot Nothing Then
                If xfrm.Offset IsNot Nothing Then
                    If xfrm.Offset.X IsNot Nothing Then x = xfrm.Offset.X.Value
                    If xfrm.Offset.Y IsNot Nothing Then y = xfrm.Offset.Y.Value
                End If
                If xfrm.Extents IsNot Nothing Then
                    If xfrm.Extents.Cx IsNot Nothing Then cx = xfrm.Extents.Cx.Value
                    If xfrm.Extents.Cy IsNot Nothing Then cy = xfrm.Extents.Cy.Value
                End If
            End If

        ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
            Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)

            ph = gf.NonVisualGraphicFrameProperties?.
                ApplicationNonVisualDrawingProperties?.
                PlaceholderShape

            name = If(
                gf.NonVisualGraphicFrameProperties?.
                    NonVisualDrawingProperties?.
                    Name?.
                    Value,
                String.Empty)

            If gf.NonVisualGraphicFrameProperties IsNot Nothing AndAlso
               gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties IsNot Nothing AndAlso
               gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id IsNot Nothing Then
                shapeId = gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id.Value
            End If

            kind = "graphicFrame"

            Dim xfrm = gf.Transform
            If xfrm IsNot Nothing Then
                If xfrm.Offset IsNot Nothing Then
                    If xfrm.Offset.X IsNot Nothing Then x = xfrm.Offset.X.Value
                    If xfrm.Offset.Y IsNot Nothing Then y = xfrm.Offset.Y.Value
                End If
                If xfrm.Extents IsNot Nothing Then
                    If xfrm.Extents.Cx IsNot Nothing Then cx = xfrm.Extents.Cx.Value
                    If xfrm.Extents.Cy IsNot Nothing Then cy = xfrm.Extents.Cy.Value
                End If
            End If
        End If

        If ph Is Nothing Then Return

        Dim textValue As String = String.Empty
        If includeText Then
            textValue = GetTextFromElementForJson(child)
        End If

        Dim area As Nullable(Of Long) = Nothing
        If cx.HasValue AndAlso cy.HasValue Then
            Dim rawArea As Double = CDbl(cx.Value) * CDbl(cy.Value)
            area = If(rawArea > Long.MaxValue, Long.MaxValue, CLng(rawArea))
        End If

        placeholders.Add(New PlaceholderJson With {
            .ShapeId = shapeId,
            .Kind = kind,
            .Name = name,
            .PlaceholderType = If(ph.Type IsNot Nothing, ph.Type.Value.ToString(), String.Empty),
            .Index = If(ph.Index IsNot Nothing, CType(ph.Index.Value, Nullable(Of UInteger)), Nothing),
            .Role = ResolvePlaceholderRoleForJson(ph, name),
            .SemanticRole = ResolvePlaceholderSemanticRole(ph, name),
            .Text = textValue,
            .TextLength = If(String.IsNullOrWhiteSpace(textValue), 0, textValue.Trim().Length),
            .X = x,
            .Y = y,
            .Cx = cx,
            .Cy = cy,
            .Area = area,
            .SourceOrder = sourceOrder,
            .IsPrimaryBodyPlaceholder = False
        })

        sourceOrder += 1
    End Sub

    Private Shared Function GetTextFromElementForJson(
        ByVal child As DocumentFormat.OpenXml.OpenXmlElement) As String

        If child Is Nothing Then Return String.Empty

        If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
            Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                CType(child, DocumentFormat.OpenXml.Presentation.Shape)

            If shp.TextBody IsNot Nothing Then
                Return ExtractTextFromTextContainer(shp.TextBody)
            End If
        ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
            Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)

            Dim tbl As DocumentFormat.OpenXml.Drawing.Table =
                gf.Graphic?.
                   GraphicData?.
                   GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Table)()

            If tbl IsNot Nothing Then
                Dim content As New System.Collections.Generic.List(Of System.String)
                ExtractTextFromTable(tbl, content)
                Return System.String.Join(vbCrLf, content).Trim()
            End If
        End If

        Return String.Empty
    End Function

    Private Shared Function ResolvePlaceholderRoleForJson(
        ByVal ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape,
        ByVal name As String) As String

        If ph Is Nothing Then Return "other"

        If ph.Type IsNot Nothing Then
            Select Case ph.Type.Value
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle
                    Return "title"

                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                    Return "subtitle"

                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object
                    Return "body"

                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.Header
                    Return "metadata"

                Case Else
                    Return "other"
            End Select
        End If

        If ph.Index IsNot Nothing Then
            If ph.Index.Value = 0UI Then Return "title"

            If ph.Index.Value = 1UI Then
                If Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.IndexOf("subtitle", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return "subtitle"
                End If

                Return "body"
            End If

            If ph.Index.Value >= 2UI Then
                Return "body"
            End If
        End If

        If Not String.IsNullOrWhiteSpace(name) Then
            If name.IndexOf("subtitle", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "subtitle"
            End If

            If name.IndexOf("title", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "title"
            End If

            If name.IndexOf("content", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               name.IndexOf("body", StringComparison.OrdinalIgnoreCase) >= 0 Then
                Return "body"
            End If
        End If

        Return "other"
    End Function

    Private Shared Function ResolvePlaceholderSemanticRole(
        ByVal ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape,
        ByVal name As System.String) As System.String

        Dim n As System.String = If(name, System.String.Empty).Trim().ToLowerInvariant()

        If ph IsNot Nothing AndAlso ph.Type IsNot Nothing Then
            Select Case ph.Type.Value
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                     DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle
                    Return "title"
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                    Return "subtitle"
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime
                    Return "date"
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer
                    Return "footer"
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber
                    Return "slide_number"
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Picture
                    Return "hero_image"
            End Select
        End If

        If n.Contains("untertitel") OrElse n.Contains("subtitle") OrElse n.Contains("sub-title") Then Return "subtitle"
        If n.Contains("übertitel") OrElse n.Contains("uebertitel") OrElse n.Contains("eyebrow") OrElse n.Contains("kicker") Then Return "eyebrow"
        If n.Contains("speaker") OrElse n.Contains("sprecher") OrElse n.Contains("presenter") OrElse n.Contains("referent") OrElse n.Contains("autor") OrElse n.Contains("author") Then Return "presenter"
        If n.Contains("position") OrElse n.Contains("funktion") OrElse n.Contains("function") OrElse n.Contains("role") OrElse n.Contains("titel funktion") Then Return "presenter_role"
        If n = "ort" OrElse n.Contains("location") OrElse n.Contains("place") OrElse n.Contains("standort") OrElse n.Contains("venue") Then Return "location"
        If n.Contains("datum") OrElse n.Contains("date") Then Return "date"
        If n.Contains("bild") OrElse n.Contains("picture") OrElse n.Contains("image") OrElse n.Contains("photo") Then Return "hero_image"
        If n.Contains("logo") Then Return "logo"
        If n.Contains("abschluss") OrElse n.Contains("closing") Then Return "closing_message"
        If n.Contains("fuss") OrElse n.Contains("footer") Then Return "footer"
        If n.Contains("nummer") OrElse n.Contains("slide number") Then Return "slide_number"
        If n.Contains("inhalt") OrElse n.Contains("content") OrElse n.Contains("body") Then Return "body"

        Return ResolvePlaceholderRoleForJson(ph, name)
    End Function

    ''' <summary>
    ''' Resolves the semantic role of a slide placeholder by also consulting its layout placeholder.
    ''' PowerPoint frequently renames slide instances to generic names such as "Text Placeholder 3",
    ''' while the corresponding layout placeholder retains the meaningful name ("Untertitel", "Ort", "Logo", etc.).
    ''' </summary>
    Private Function ResolveEffectivePlaceholderSemanticRole(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal shape As DocumentFormat.OpenXml.Presentation.Shape) As System.String

        If shape Is Nothing Then Return "other"

        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        Dim ownName As System.String =
            If(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)

        Dim directRole As System.String = ResolvePlaceholderSemanticRole(ph, ownName)
        If ph Is Nothing OrElse sp?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree Is Nothing Then
            Return directRole
        End If

        Dim layoutShapes As System.Collections.Generic.IEnumerable(Of DocumentFormat.OpenXml.Presentation.Shape) =
            sp.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

        Dim layoutMatch As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        If ph.Index IsNot Nothing Then
            layoutMatch = layoutShapes.FirstOrDefault(
                Function(candidate)
                    Dim cph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                        candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    Return cph?.Index IsNot Nothing AndAlso cph.Index.Value = ph.Index.Value
                End Function)
        End If

        If layoutMatch Is Nothing AndAlso ph.Type IsNot Nothing Then
            Dim sameType As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
                layoutShapes.
                    Where(
                        Function(candidate)
                            Dim cph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                                candidate.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                            Return cph?.Type IsNot Nothing AndAlso cph.Type.Value = ph.Type.Value
                        End Function).
                    ToList()

            If sameType.Count = 1 Then layoutMatch = sameType(0)
        End If

        If layoutMatch Is Nothing Then Return directRole

        Dim layoutPh As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            layoutMatch.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        Dim layoutName As System.String =
            If(layoutMatch.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
        Dim layoutRole As System.String = ResolvePlaceholderSemanticRole(layoutPh, layoutName)

        ' Prefer the layout's semantic information whenever the slide instance only exposes a generic
        ' body/other role. This is critical for corporate templates that use Body placeholders for
        ' subtitles, locations, logos, eyebrow strips, presenter fields, etc.
        If Not System.String.IsNullOrWhiteSpace(layoutRole) AndAlso
           Not System.String.Equals(layoutRole, "other", System.StringComparison.OrdinalIgnoreCase) Then

            If System.String.Equals(directRole, "body", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(directRole, "other", System.StringComparison.OrdinalIgnoreCase) Then
                Return layoutRole
            End If
        End If

        Return directRole
    End Function

    ''' <summary>
    ''' Returns True when a placeholder shape contains relationship-backed visual artwork (for example
    ''' an SVG/logo stored in a:blipFill). Such shapes must stay on the layout and must not be cloned
    ''' as slide-level placeholder overrides unless their image relationships are also copied.
    ''' </summary>
    Private Shared Function ShapeContainsRelationshipBackedVisual(
        ByVal shape As DocumentFormat.OpenXml.Presentation.Shape) As System.Boolean

        If shape?.ShapeProperties Is Nothing Then Return False
        Return shape.ShapeProperties.Descendants(Of DocumentFormat.OpenXml.Drawing.Blip)().Any()
    End Function

    Private Shared Sub MarkPrimaryBodyPlaceholder(
        ByVal placeholders As System.Collections.Generic.List(Of PlaceholderJson),
        ByVal preferText As System.Boolean)

        If placeholders Is Nothing OrElse placeholders.Count = 0 Then Return

        For Each ph In placeholders
            ph.IsPrimaryBodyPlaceholder = False
        Next

        Dim bodies = placeholders.
            Where(Function(p) String.Equals(p.Role, "body", StringComparison.OrdinalIgnoreCase)).
            ToList()

        If bodies.Count = 0 Then Return

        Dim primary As PlaceholderJson = Nothing

        If preferText Then
            primary = bodies.
                OrderByDescending(Function(p) p.TextLength).
                ThenByDescending(Function(p) If(p.Area.HasValue, p.Area.Value, 0L)).
                ThenBy(Function(p) p.SourceOrder).
                FirstOrDefault()
        Else
            primary = bodies.
                OrderByDescending(Function(p) If(p.Area.HasValue, p.Area.Value, 0L)).
                ThenBy(Function(p) If(p.Index.HasValue, CLng(p.Index.Value), Long.MaxValue)).
                ThenBy(Function(p) p.SourceOrder).
                FirstOrDefault()
        End If

        If primary IsNot Nothing Then
            primary.IsPrimaryBodyPlaceholder = True
        End If
    End Sub


    ''' <summary>
    ''' Collects placeholder type names from a shape tree.
    ''' </summary>
    ''' <param name="tree">The shape tree to scan.</param>
    ''' <param name="placeholders">List to populate with placeholder type names.</param>
    Private Shared Sub CollectPlaceholdersFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal placeholders As System.Collections.Generic.List(Of System.String))

        If tree Is Nothing Then Return

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements

            If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(child, DocumentFormat.OpenXml.Presentation.Shape)

                Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    shp.NonVisualShapeProperties

                If nv IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type IsNot Nothing Then

                    placeholders.Add(
                        nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type.Value.ToString())
                End If

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                ' Recursively process children within groups
                Dim grp As DocumentFormat.OpenXml.Presentation.GroupShape =
                    CType(child, DocumentFormat.OpenXml.Presentation.GroupShape)

                If grp.ChildElements.Count > 0 Then
                    CollectPlaceholdersFromGroup(grp, placeholders)
                End If
            End If
            ' Note: GraphicFrame (tables/charts) do not have placeholders
        Next
    End Sub

    ''' <summary>
    ''' Recursively collects placeholder types from a group shape.
    ''' </summary>
    ''' <param name="group">The group shape to scan.</param>
    ''' <param name="placeholders">List to populate with placeholder type names.</param>
    Private Shared Sub CollectPlaceholdersFromGroup(
        ByVal group As DocumentFormat.OpenXml.Presentation.GroupShape,
        ByVal placeholders As System.Collections.Generic.List(Of System.String))

        For Each inner As DocumentFormat.OpenXml.OpenXmlElement In group.ChildElements
            If TypeOf inner Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(inner, DocumentFormat.OpenXml.Presentation.Shape)

                Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    shp.NonVisualShapeProperties

                If nv IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape IsNot Nothing AndAlso
                   nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type IsNot Nothing Then

                    placeholders.Add(
                        nv.ApplicationNonVisualDrawingProperties.PlaceholderShape.Type.Value.ToString())
                End If

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectPlaceholdersFromGroup(
                    CType(inner, DocumentFormat.OpenXml.Presentation.GroupShape), placeholders)
            End If
        Next
    End Sub

    ''' <summary>
    ''' Collects text content from shapes, groups, and tables in a shape tree.
    ''' </summary>
    ''' <param name="tree">The shape tree to scan.</param>
    ''' <param name="content">List to populate with extracted text.</param>
    Private Shared Sub CollectTextsFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal content As System.Collections.Generic.List(Of System.String))

        If tree Is Nothing Then Return

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements

            If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(child, DocumentFormat.OpenXml.Presentation.Shape)

                If shp.TextBody IsNot Nothing Then
                    Dim txt As System.String = ExtractTextFromTextContainer(shp.TextBody)
                    If Not System.String.IsNullOrWhiteSpace(txt) Then content.Add(txt)
                End If

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectTextsFromGroup(
                    CType(child, DocumentFormat.OpenXml.Presentation.GroupShape), content)

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                    CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)

                Dim g As DocumentFormat.OpenXml.Drawing.Graphic = gf.Graphic
                If g IsNot Nothing AndAlso g.GraphicData IsNot Nothing Then
                    Dim gd As DocumentFormat.OpenXml.Drawing.GraphicData = g.GraphicData

                    ' Extract text from tables
                    Dim tbl As DocumentFormat.OpenXml.Drawing.Table =
                        gd.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Table)()
                    If tbl IsNot Nothing Then
                        ExtractTextFromTable(tbl, content)
                    End If
                    ' Note: Charts/SmartArt could be handled here additionally
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extracts concatenated text from all paragraphs in a text container element.
    ''' </summary>
    ''' <param name="container">The container element (TextBody or similar).</param>
    ''' <returns>Concatenated text content trimmed.</returns>
    Private Shared Function ExtractTextFromTextContainer(
        ByVal container As DocumentFormat.OpenXml.OpenXmlElement) As System.String

        If container Is Nothing Then Return System.String.Empty

        Dim parts As New System.Collections.Generic.List(Of System.String)()

        ' Walk all Drawing.Paragraph descendants regardless of the exact TextBody type
        For Each p As DocumentFormat.OpenXml.Drawing.Paragraph In
            container.Descendants(Of DocumentFormat.OpenXml.Drawing.Paragraph)()

            Dim runs As New System.Collections.Generic.List(Of System.String)()

            For Each r As DocumentFormat.OpenXml.Drawing.Run In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Run)()
                If r IsNot Nothing AndAlso r.Text IsNot Nothing Then
                    runs.Add(r.Text.Text)
                End If
            Next

            For Each br As DocumentFormat.OpenXml.Drawing.Break In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Break)()
                runs.Add(vbLf)
            Next

            For Each fld As DocumentFormat.OpenXml.Drawing.Field In
                p.Elements(Of DocumentFormat.OpenXml.Drawing.Field)()
                If fld IsNot Nothing AndAlso fld.Text IsNot Nothing Then
                    runs.Add(fld.Text.Text)
                End If
            Next

            parts.Add(System.String.Join(System.String.Empty, runs))
        Next

        Return System.String.Join(vbCrLf, parts).Trim()
    End Function

    ''' <summary>
    ''' Recursively collects text content from a group shape.
    ''' </summary>
    ''' <param name="group">The group shape to scan.</param>
    ''' <param name="content">List to populate with extracted text.</param>
    Private Shared Sub CollectTextsFromGroup(
        ByVal group As DocumentFormat.OpenXml.Presentation.GroupShape,
        ByVal content As System.Collections.Generic.List(Of System.String))

        ' Cannot use IsNot comparison on OpenXmlElementList - only check Count
        If group.ChildElements.Count = 0 Then Return

        For Each inner As DocumentFormat.OpenXml.OpenXmlElement In group.ChildElements

            If TypeOf inner Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape =
                    CType(inner, DocumentFormat.OpenXml.Presentation.Shape)

                If shp.TextBody IsNot Nothing Then
                    Dim txt As System.String = ExtractTextFromTextContainer(shp.TextBody)
                    If Not System.String.IsNullOrWhiteSpace(txt) Then content.Add(txt)
                End If

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                CollectTextsFromGroup(
                    CType(inner, DocumentFormat.OpenXml.Presentation.GroupShape), content)

            ElseIf TypeOf inner Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame =
                    CType(inner, DocumentFormat.OpenXml.Presentation.GraphicFrame)

                Dim g As DocumentFormat.OpenXml.Drawing.Graphic = gf.Graphic
                If g IsNot Nothing AndAlso g.GraphicData IsNot Nothing Then
                    Dim gd As DocumentFormat.OpenXml.Drawing.GraphicData = g.GraphicData

                    Dim tbl As DocumentFormat.OpenXml.Drawing.Table =
                        gd.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Table)()
                    If tbl IsNot Nothing Then
                        ExtractTextFromTable(tbl, content)
                    End If
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Extracts text from a Drawing.Table element, adding tab-separated rows.
    ''' </summary>
    ''' <param name="table">The table element to extract text from.</param>
    ''' <param name="content">List to populate with row text.</param>
    Private Shared Sub ExtractTextFromTable(
        ByVal table As DocumentFormat.OpenXml.Drawing.Table,
        ByVal content As System.Collections.Generic.List(Of System.String))

        If table Is Nothing Then Return

        For Each row As DocumentFormat.OpenXml.Drawing.TableRow In
            table.Elements(Of DocumentFormat.OpenXml.Drawing.TableRow)()

            Dim rowTexts As New System.Collections.Generic.List(Of System.String)()

            For Each cell As DocumentFormat.OpenXml.Drawing.TableCell In
                row.Elements(Of DocumentFormat.OpenXml.Drawing.TableCell)()

                If cell IsNot Nothing AndAlso cell.TextBody IsNot Nothing Then
                    Dim cellText As System.String = ExtractTextFromTextContainer(cell.TextBody)
                    rowTexts.Add(cellText)
                End If
            Next

            Dim line As System.String = System.String.Join(vbTab, rowTexts)
            If Not System.String.IsNullOrWhiteSpace(line) Then content.Add(line)
        Next
    End Sub


#End Region

    ''' <summary>
    ''' Validates a PPTX file by checking ZIP archive integrity and entry accessibility.
    ''' </summary>
    ''' <param name="path">Full path to the PowerPoint file.</param>
    ''' <returns>True if the file is a valid, readable PPTX package.</returns>
    Private Function IsValidPptxPackage(path As String) As Boolean
        Try
            Using archive As New System.IO.Compression.ZipArchive(System.IO.File.OpenRead(path), IO.Compression.ZipArchiveMode.Read)
                For Each entry In archive.Entries
                    ' Optional: basic sanity check — size not huge, name not empty
                    If String.IsNullOrWhiteSpace(entry.FullName) Then Return False
                    ' Read small text files fully to catch unreadable/corrupt XML
                    If entry.Length > 0 AndAlso entry.Length < 5_000_000 Then
                        Using s = entry.Open()
                            ' Read a few bytes to ensure stream is accessible
                            Dim buffer(255) As Byte
                            Dim read = s.Read(buffer, 0, buffer.Length)
                        End Using
                    End If
                Next
            End Using
            Return True
        Catch
            ' If ZIP open fails or reading any entry fails, it's not valid
            Return False
        End Try
    End Function

#Region "JSON DTOs"

    ''' <summary>
    ''' Represents a single placeholder with detailed metadata for JSON serialization.
    ''' This is additive and does not replace the legacy slide-level placeholders/content arrays.
    ''' </summary>
    Public Class PlaceholderJson
        <JsonPropertyName("shapeId")>
        Public Property ShapeId As Nullable(Of UInteger)

        <JsonPropertyName("kind")>
        Public Property Kind As String

        <JsonPropertyName("name")>
        Public Property Name As String

        <JsonPropertyName("type")>
        Public Property PlaceholderType As String

        <JsonPropertyName("index")>
        Public Property Index As Nullable(Of UInteger)

        <JsonPropertyName("role")>
        Public Property Role As String

        <JsonPropertyName("semanticRole")>
        Public Property SemanticRole As System.String

        <JsonPropertyName("text")>
        Public Property Text As String

        <JsonPropertyName("textLength")>
        Public Property TextLength As Integer

        <JsonPropertyName("x")>
        Public Property X As Nullable(Of Long)

        <JsonPropertyName("y")>
        Public Property Y As Nullable(Of Long)

        <JsonPropertyName("cx")>
        Public Property Cx As Nullable(Of Long)

        <JsonPropertyName("cy")>
        Public Property Cy As Nullable(Of Long)

        <JsonPropertyName("area")>
        Public Property Area As Nullable(Of Long)

        <JsonPropertyName("sourceOrder")>
        Public Property SourceOrder As Integer

        <JsonPropertyName("isPrimaryBodyPlaceholder")>
        Public Property IsPrimaryBodyPlaceholder As Boolean

        <JsonPropertyName("effectiveX")>
        Public Property EffectiveX As Nullable(Of Long)
        <JsonPropertyName("effectiveY")>
        Public Property EffectiveY As Nullable(Of Long)
        <JsonPropertyName("effectiveCx")>
        Public Property EffectiveCx As Nullable(Of Long)
        <JsonPropertyName("effectiveCy")>
        Public Property EffectiveCy As Nullable(Of Long)
        <JsonPropertyName("textStyle")>
        Public Property TextStyle As TextStyleJson
    End Class

    ''' <summary>
    ''' Represents a single slide's metadata and content for JSON serialization.
    ''' </summary>
    Public Class SlideJson
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String

        <JsonPropertyName("slideId")>
        Public Property SlideId As UInteger

        <JsonPropertyName("index")>
        Public Property Index As Integer

        <JsonPropertyName("title")>
        Public Property Title As String

        <JsonPropertyName("layout")>
        Public Property Layout As String

        <JsonPropertyName("master")>
        Public Property Master As String

        <JsonPropertyName("placeholders")>
        Public Property Placeholders As List(Of String)

        <JsonPropertyName("content")>
        Public Property Content As List(Of String)

        <JsonPropertyName("placeholderDetails")>
        Public Property PlaceholderDetails As List(Of PlaceholderJson)

        <JsonPropertyName("layoutId")>
        Public Property LayoutId As String
        <JsonPropertyName("layoutRelId")>
        Public Property LayoutRelId As String
        <JsonPropertyName("masterId")>
        Public Property MasterId As String
        <JsonPropertyName("visualElements")>
        Public Property VisualElements As List(Of VisualElementJson)
        <JsonPropertyName("editableSlots")>
        Public Property EditableSlots As List(Of EditableSlotJson)
        <JsonPropertyName("cloneAsTemplateScore")>
        Public Property CloneAsTemplateScore As Integer
        <JsonPropertyName("cloneAsTemplateReason")>
        Public Property CloneAsTemplateReason As String
        <JsonPropertyName("backgroundColor")>
        Public Property BackgroundColor As String
    End Class

    ''' <summary>
    ''' Represents a slide layout's metadata for JSON serialization.
    ''' </summary>
    Public Class LayoutJson
        <JsonPropertyName("name")>
        Public Property Name As String

        <JsonPropertyName("layoutId")>
        Public Property LayoutId As String

        <JsonPropertyName("layoutRelId")>
        Public Property LayoutRelId As String

        <JsonPropertyName("master")>
        Public Property Master As String

        <JsonPropertyName("placeholderDetails")>
        Public Property PlaceholderDetails As List(Of PlaceholderJson)

        <JsonPropertyName("masterId")>
        Public Property MasterId As String
        <JsonPropertyName("signature")>
        Public Property Signature As String
        <JsonPropertyName("semanticRole")>
        Public Property SemanticRole As System.String
        <JsonPropertyName("recommendedUses")>
        Public Property RecommendedUses As List(Of String)
        <JsonPropertyName("avoidUses")>
        Public Property AvoidUses As List(Of String)
        <JsonPropertyName("usageCount")>
        Public Property UsageCount As Integer
        <JsonPropertyName("exampleSlideKeys")>
        Public Property ExampleSlideKeys As List(Of String)
        <JsonPropertyName("qualityScore")>
        Public Property QualityScore As Integer
        <JsonPropertyName("safeForComponents")>
        Public Property SafeForComponents As Boolean
        <JsonPropertyName("backgroundColor")>
        Public Property BackgroundColor As String
        <JsonPropertyName("visualElements")>
        Public Property VisualElements As List(Of VisualElementJson)
    End Class

    ''' <summary>
    ''' Represents slide dimensions in EMUs for JSON serialization.
    ''' </summary>
    Public Class SlideSizeJson
        <JsonPropertyName("width")>
        Public Property Width As Long

        <JsonPropertyName("height")>
        Public Property Height As Long
    End Class

    ''' <summary>
    ''' Root DTO for presentation JSON export including slides and layouts.
    ''' </summary>
    Public Class PresentationJson
        <JsonPropertyName("version")>
        Public Property Version As String
        <JsonPropertyName("title")>
        Public Property Title As String

        <JsonPropertyName("slideSize")>
        Public Property SlideSize As SlideSizeJson

        <JsonPropertyName("slides")>
        Public Property Slides As List(Of SlideJson)

        <JsonPropertyName("layouts")>
        Public Property Layouts As List(Of LayoutJson)
        <JsonPropertyName("designProfile")>
        Public Property DesignProfile As DesignProfileJson
    End Class


    Public Class TextStyleJson
        <JsonPropertyName("fontFamily")>
        Public Property FontFamily As String
        <JsonPropertyName("fontSize")>
        Public Property FontSize As Nullable(Of Double)
        <JsonPropertyName("bold")>
        Public Property Bold As Nullable(Of Boolean)
        <JsonPropertyName("italic")>
        Public Property Italic As Nullable(Of Boolean)
        <JsonPropertyName("color")>
        Public Property Color As String
        <JsonPropertyName("colorRef")>
        Public Property ColorRef As String
        <JsonPropertyName("alignment")>
        Public Property Alignment As String
        <JsonPropertyName("bullet")>
        Public Property Bullet As Nullable(Of Boolean)
        <JsonPropertyName("bulletChar")>
        Public Property BulletChar As String
    End Class

    Public Class VisualElementJson
        <JsonPropertyName("shapeId")>
        Public Property ShapeId As Nullable(Of UInteger)
        <JsonPropertyName("kind")>
        Public Property Kind As String
        <JsonPropertyName("name")>
        Public Property Name As String
        <JsonPropertyName("shapeType")>
        Public Property ShapeType As String
        <JsonPropertyName("x")>
        Public Property X As Nullable(Of Long)
        <JsonPropertyName("y")>
        Public Property Y As Nullable(Of Long)
        <JsonPropertyName("cx")>
        Public Property Cx As Nullable(Of Long)
        <JsonPropertyName("cy")>
        Public Property Cy As Nullable(Of Long)
        <JsonPropertyName("fillColor")>
        Public Property FillColor As String
        <JsonPropertyName("outlineColor")>
        Public Property OutlineColor As String
        <JsonPropertyName("text")>
        Public Property Text As String
        <JsonPropertyName("isPlaceholder")>
        Public Property IsPlaceholder As Boolean
    End Class

    Public Class EditableSlotJson
        <JsonPropertyName("shapeId")>
        Public Property ShapeId As UInteger
        <JsonPropertyName("name")>
        Public Property Name As String
        <JsonPropertyName("role")>
        Public Property Role As String
        <JsonPropertyName("semanticRole")>
        Public Property SemanticRole As System.String
        <JsonPropertyName("placeholderType")>
        Public Property PlaceholderType As String
        <JsonPropertyName("placeholderIndex")>
        Public Property PlaceholderIndex As Nullable(Of UInteger)
        <JsonPropertyName("text")>
        Public Property Text As String
        <JsonPropertyName("x")>
        Public Property X As Nullable(Of Long)
        <JsonPropertyName("y")>
        Public Property Y As Nullable(Of Long)
        <JsonPropertyName("cx")>
        Public Property Cx As Nullable(Of Long)
        <JsonPropertyName("cy")>
        Public Property Cy As Nullable(Of Long)
        <JsonPropertyName("textStyle")>
        Public Property TextStyle As TextStyleJson
    End Class

    Public Class DesignProfileJson
        <JsonPropertyName("templateStrength")>
        Public Property TemplateStrength As Integer
        <JsonPropertyName("modeHint")>
        Public Property ModeHint As String
        <JsonPropertyName("headingFont")>
        Public Property HeadingFont As String
        <JsonPropertyName("bodyFont")>
        Public Property BodyFont As String
        <JsonPropertyName("palette")>
        Public Property Palette As System.Collections.Generic.Dictionary(Of String, String)
        <JsonPropertyName("observedColors")>
        Public Property ObservedColors As List(Of String)
        <JsonPropertyName("commonShapeTypes")>
        Public Property CommonShapeTypes As List(Of String)
        <JsonPropertyName("visualDensity")>
        Public Property VisualDensity As String
        <JsonPropertyName("existingSlideCount")>
        Public Property ExistingSlideCount As Integer
        <JsonPropertyName("sampleDeckLikelihood")>
        Public Property SampleDeckLikelihood As System.String
        <JsonPropertyName("sampleDeckConfidence")>
        Public Property SampleDeckConfidence As System.Int32
        <JsonPropertyName("sampleSlideKeys")>
        Public Property SampleSlideKeys As System.Collections.Generic.List(Of System.String)
        <JsonPropertyName("sampleDeckReason")>
        Public Property SampleDeckReason As System.String
        <JsonPropertyName("guidance")>
        Public Property Guidance As String
        <JsonPropertyName("masterBackgroundColor")>
        Public Property MasterBackgroundColor As String
        <JsonPropertyName("masterVisualElements")>
        Public Property MasterVisualElements As List(Of VisualElementJson)
    End Class

    ''' <summary>
    ''' Internal class to track placeholder types present in a layout.
    ''' </summary>
    Private NotInheritable Class LayoutInfo
        Public Property HasTitle As System.Boolean
        Public Property HasCenteredTitle As System.Boolean
        Public Property HasSubTitle As System.Boolean
        Public Property HasBody As System.Boolean
    End Class

    ''' <summary>
    ''' Base class for action plan operations.
    ''' </summary>
    Public MustInherit Class ActionBase
        <JsonPropertyName("op")>
        Public Property Op As String
    End Class

    ''' <summary>
    ''' Represents anchor positioning for slide insertion.
    ''' </summary>
    Public Class Anchor
        <JsonPropertyName("mode")>
        Public Property Mode As String
        <JsonPropertyName("by")>
        Public Property By As AnchorBy
    End Class

    ''' <summary>
    ''' Specifies the slide key for anchor positioning.
    ''' </summary>
    Public Class AnchorBy
        <JsonPropertyName("slideKey")>
        Public Property SlideKey As String
    End Class

    ''' <summary>
    ''' Represents an add_slide action with anchor, layout, and elements.
    ''' </summary>
    Public Class AddSlideAction
        Inherits ActionBase
        <JsonPropertyName("anchor")> Public Property Anchor As Anchor
        <JsonPropertyName("layoutRelId")> Public Property LayoutRelId As String
        <JsonPropertyName("elements")> Public Property Elements As List(Of JsonElement)
    End Class


    ''' <summary>
    ''' Index mapping slide keys to IDs and IDs to positions.
    ''' </summary>
    Public Class DeckIndex
        Public Property SlideKeyById As Dictionary(Of String, UInteger)
        Public Property IndexBySlideId As Dictionary(Of UInteger, Integer)
    End Class

#End Region


#Region "Enhanced Design Metadata"

    Private Sub EnrichPresentationJson(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal result As PresentationJson)

        If presPart Is Nothing OrElse result Is Nothing Then Return

        ' Enrichment is advisory metadata for the LLM. A malformed or unusual
        ' theme/master must never make the core presentation extraction fail.
        Try
            result.Version = "2.4.0"

            Dim slidePartById As New System.Collections.Generic.Dictionary(Of UInteger, DocumentFormat.OpenXml.Packaging.SlidePart)()
            Dim slideIdList As DocumentFormat.OpenXml.Presentation.SlideIdList = presPart.Presentation?.SlideIdList
            If slideIdList IsNot Nothing Then
                For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In slideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                    If sid.RelationshipId Is Nothing Then Continue For
                    Try
                        Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                        If sp IsNot Nothing Then slidePartById(sid.Id.Value) = sp
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("Could not enrich slide metadata: " & ex.Message)
                    End Try
                Next
            End If

            If result.Slides IsNot Nothing Then
                For Each slideMeta As SlideJson In result.Slides
                    If slideMeta Is Nothing OrElse Not slidePartById.ContainsKey(slideMeta.SlideId) Then Continue For
                    Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = slidePartById(slideMeta.SlideId)
                    Dim lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp.SlideLayoutPart
                    If lp IsNot Nothing Then
                        slideMeta.LayoutId = If(lp.Uri IsNot Nothing, lp.Uri.ToString(), System.String.Empty)
                        slideMeta.MasterId = If(lp.SlideMasterPart?.Uri IsNot Nothing, lp.SlideMasterPart.Uri.ToString(), System.String.Empty)
                        If lp.SlideMasterPart IsNot Nothing Then
                            Try
                                slideMeta.LayoutRelId = lp.SlideMasterPart.GetIdOfPart(lp)
                            Catch ex As System.Exception
                                slideMeta.LayoutRelId = System.String.Empty
                            End Try
                        End If
                    End If

                    slideMeta.VisualElements = CollectVisualElementsForJson(sp)
                    slideMeta.BackgroundColor = ExtractEffectiveBackgroundColor(sp)
                    slideMeta.EditableSlots = CollectEditableSlotsForJson(sp)
                    Dim reason As System.String = System.String.Empty
                    slideMeta.CloneAsTemplateScore = ScoreSlideAsReusableTemplate(sp, slideMeta, reason)
                    slideMeta.CloneAsTemplateReason = reason

                    If slideMeta.PlaceholderDetails IsNot Nothing Then
                        For Each phMeta As PlaceholderJson In slideMeta.PlaceholderDetails
                            PopulateEffectivePlaceholderMetadata(sp, phMeta)
                        Next
                    End If
                Next
            End If

            If result.Layouts IsNot Nothing Then
                For Each layoutMeta As LayoutJson In result.Layouts
                    If layoutMeta Is Nothing Then Continue For
                    Dim lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = FindLayoutByUriOrName(presPart, layoutMeta.LayoutId, layoutMeta.Name, layoutMeta.Master)
                    If lp Is Nothing Then Continue For

                    layoutMeta.MasterId = If(lp.SlideMasterPart?.Uri IsNot Nothing, lp.SlideMasterPart.Uri.ToString(), System.String.Empty)
                    layoutMeta.Signature = BuildLayoutSignature(lp)
                    layoutMeta.SemanticRole = ClassifyLayoutSemanticRole(lp)
                    layoutMeta.BackgroundColor = ExtractBackgroundColor(lp.SlideLayout?.CommonSlideData?.Background)
                    layoutMeta.VisualElements = CollectVisualElementsFromShapeTree(lp.SlideLayout?.CommonSlideData?.ShapeTree)
                    layoutMeta.RecommendedUses = GetRecommendedUsesForLayout(layoutMeta.SemanticRole)
                    layoutMeta.AvoidUses = GetAvoidUsesForLayout(layoutMeta.SemanticRole)
                    layoutMeta.SafeForComponents = IsLayoutSafeForComponents(lp)

                    If layoutMeta.PlaceholderDetails IsNot Nothing Then
                        For Each phMeta As PlaceholderJson In layoutMeta.PlaceholderDetails
                            PopulateEffectivePlaceholderMetadata(lp, phMeta)
                        Next
                    End If

                    Dim examples As New System.Collections.Generic.List(Of System.String)()
                    Dim usageCount As System.Int32 = 0
                    If result.Slides IsNot Nothing Then
                        For Each slideMeta As SlideJson In result.Slides
                            If slideMeta Is Nothing Then Continue For
                            If System.String.Equals(slideMeta.LayoutId, layoutMeta.LayoutId, System.StringComparison.OrdinalIgnoreCase) Then
                                usageCount += 1
                                If examples.Count < 3 AndAlso Not System.String.IsNullOrWhiteSpace(slideMeta.SlideKey) Then examples.Add(slideMeta.SlideKey)
                            End If
                        Next
                    End If
                    layoutMeta.UsageCount = usageCount
                    layoutMeta.ExampleSlideKeys = examples
                    layoutMeta.QualityScore = ScoreLayoutQuality(lp, usageCount)
                Next
            End If

            result.DesignProfile = BuildDesignProfile(presPart, result)
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Enhanced presentation metadata was partially skipped: " & ex.ToString())
            ' Keep all core slides/layouts already collected. The LLM can still
            ' operate with the base metadata when optional design enrichment fails.
        End Try
    End Sub

    Private Function FindLayoutByUriOrName(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal uri As System.String,
        ByVal name As System.String,
        ByVal masterName As System.String) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        If presPart Is Nothing Then Return Nothing
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            Dim currentMasterName As System.String = GetMasterName(sm)
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                If Not System.String.IsNullOrWhiteSpace(uri) AndAlso lp.Uri IsNot Nothing AndAlso
                   System.String.Equals(lp.Uri.ToString(), uri, System.StringComparison.OrdinalIgnoreCase) Then Return lp

                If Not System.String.IsNullOrWhiteSpace(name) AndAlso
                   System.String.Equals(GetLayoutName(lp), name, System.StringComparison.OrdinalIgnoreCase) AndAlso
                   (System.String.IsNullOrWhiteSpace(masterName) OrElse System.String.Equals(currentMasterName, masterName, System.StringComparison.OrdinalIgnoreCase)) Then Return lp
            Next
        Next
        Return Nothing
    End Function

    Private Function CollectVisualElementsForJson(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As System.Collections.Generic.List(Of VisualElementJson)

        Return CollectVisualElementsFromShapeTree(sp?.Slide?.CommonSlideData?.ShapeTree)
    End Function

    Private Function CollectVisualElementsFromShapeTree(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree) As System.Collections.Generic.List(Of VisualElementJson)

        Dim answer As New System.Collections.Generic.List(Of VisualElementJson)()
        If tree Is Nothing Then Return answer

        For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements
            If answer.Count >= 36 Then Exit For

            If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                Dim shp As DocumentFormat.OpenXml.Presentation.Shape = CType(child, DocumentFormat.OpenXml.Presentation.Shape)
                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                Dim v As New VisualElementJson With {
                    .ShapeId = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id IsNot Nothing, CType(shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value, Nullable(Of UInteger)), Nothing),
                    .Kind = "shape",
                    .Name = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty),
                    .IsPlaceholder = (ph IsNot Nothing),
                    .Text = TruncateMetadataText(If(shp.TextBody IsNot Nothing, ExtractTextFromTextContainer(shp.TextBody), System.String.Empty), 180),
                    .FillColor = ExtractSolidFillColor(shp.ShapeProperties),
                    .OutlineColor = ExtractOutlineColor(shp.ShapeProperties)
                }
                Dim geom As DocumentFormat.OpenXml.Drawing.PresetGeometry = shp.ShapeProperties?.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.PresetGeometry)()
                If geom?.Preset IsNot Nothing Then v.ShapeType = geom.Preset.Value.ToString()
                PopulateVisualTransform(v, shp.ShapeProperties?.Transform2D)
                If Not v.IsPlaceholder OrElse Not System.String.IsNullOrWhiteSpace(v.Text) Then answer.Add(v)

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.Picture Then
                Dim pic As DocumentFormat.OpenXml.Presentation.Picture = CType(child, DocumentFormat.OpenXml.Presentation.Picture)
                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                Dim v As New VisualElementJson With {
                    .ShapeId = If(pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Id IsNot Nothing, CType(pic.NonVisualPictureProperties.NonVisualDrawingProperties.Id.Value, Nullable(Of UInteger)), Nothing),
                    .Kind = "picture",
                    .Name = If(pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty),
                    .IsPlaceholder = (ph IsNot Nothing)
                }
                PopulateVisualTransform(v, pic.ShapeProperties?.Transform2D)
                answer.Add(v)

            ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame = CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)
                Dim kind As System.String = "graphicFrame"
                Dim uri As System.String = If(gf.Graphic?.GraphicData?.Uri?.Value, System.String.Empty)
                If uri.IndexOf("table", System.StringComparison.OrdinalIgnoreCase) >= 0 Then kind = "table"
                If uri.IndexOf("chart", System.StringComparison.OrdinalIgnoreCase) >= 0 Then kind = "chart"
                If uri.IndexOf("diagram", System.StringComparison.OrdinalIgnoreCase) >= 0 Then kind = "diagram"
                Dim v As New VisualElementJson With {
                    .ShapeId = If(gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Id IsNot Nothing, CType(gf.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id.Value, Nullable(Of UInteger)), Nothing),
                    .Kind = kind,
                    .Name = If(gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty),
                    .IsPlaceholder = (gf.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape IsNot Nothing)
                }
                If gf.Transform IsNot Nothing Then
                    If gf.Transform.Offset IsNot Nothing Then
                        v.X = gf.Transform.Offset.X?.Value
                        v.Y = gf.Transform.Offset.Y?.Value
                    End If
                    If gf.Transform.Extents IsNot Nothing Then
                        v.Cx = gf.Transform.Extents.Cx?.Value
                        v.Cy = gf.Transform.Extents.Cy?.Value
                    End If
                End If
                answer.Add(v)
            End If
        Next

        Return answer
    End Function

    Private Function CollectEditableSlotsForJson(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As System.Collections.Generic.List(Of EditableSlotJson)

        Dim answer As New System.Collections.Generic.List(Of EditableSlotJson)()
        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp?.Slide?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return answer

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            If shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id Is Nothing Then Continue For
            If shp.TextBody Is Nothing Then Continue For

            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            Dim role As System.String = If(ph IsNot Nothing, ResolvePlaceholderRoleForJson(ph, If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)), "freeText")
            Dim slot As New EditableSlotJson With {
                .ShapeId = shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value,
                .Name = If(shp.NonVisualShapeProperties.NonVisualDrawingProperties.Name?.Value, System.String.Empty),
                .Role = role,
                .SemanticRole = If(ph IsNot Nothing, ResolveEffectivePlaceholderSemanticRole(sp, shp), "freeText"),
                .PlaceholderType = If(ph?.Type IsNot Nothing, ph.Type.Value.ToString(), System.String.Empty),
                .PlaceholderIndex = If(ph?.Index IsNot Nothing, CType(ph.Index.Value, Nullable(Of UInteger)), Nothing),
                .Text = TruncateMetadataText(ExtractTextFromTextContainer(shp.TextBody), 240),
                .TextStyle = ExtractEffectiveTextStyleForShape(sp, shp)
            }

            Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shp)
            If xfrm IsNot Nothing Then
                If xfrm.Offset IsNot Nothing Then
                    slot.X = xfrm.Offset.X?.Value
                    slot.Y = xfrm.Offset.Y?.Value
                End If
                If xfrm.Extents IsNot Nothing Then
                    slot.Cx = xfrm.Extents.Cx?.Value
                    slot.Cy = xfrm.Extents.Cy?.Value
                End If
            End If
            ResolveTextStyleColorForJson(slot.TextStyle, sp?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme)
            answer.Add(slot)
        Next

        Return answer
    End Function

    Private Sub ResolveTextStyleColorForJson(
        ByVal style As TextStyleJson,
        ByVal theme As DocumentFormat.OpenXml.Drawing.Theme)

        If style Is Nothing OrElse System.String.IsNullOrWhiteSpace(style.Color) Then Return
        Dim raw As System.String = style.Color
        style.ColorRef = raw
        If raw.StartsWith("scheme:", System.StringComparison.OrdinalIgnoreCase) Then
            style.Color = ResolveThemeColorReference(theme, raw, "#0F172A")
        End If
    End Sub

    Private Function TruncateMetadataText(ByVal value As System.String, ByVal maxLength As System.Int32) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return System.String.Empty
        Dim normalized As System.String = value.Replace(vbCrLf, " | ").Replace(vbLf, " | ").Trim()
        If normalized.Length <= maxLength Then Return normalized
        Return normalized.Substring(0, System.Math.Max(0, maxLength - 1)).TrimEnd() & "…"
    End Function

    Private Sub PopulateVisualTransform(ByVal target As VisualElementJson, ByVal xfrm As DocumentFormat.OpenXml.Drawing.Transform2D)
        If target Is Nothing OrElse xfrm Is Nothing Then Return
        If xfrm.Offset IsNot Nothing Then
            target.X = xfrm.Offset.X?.Value
            target.Y = xfrm.Offset.Y?.Value
        End If
        If xfrm.Extents IsNot Nothing Then
            target.Cx = xfrm.Extents.Cx?.Value
            target.Cy = xfrm.Extents.Cy?.Value
        End If
    End Sub

    Private Function ExtractSolidFillColor(ByVal element As DocumentFormat.OpenXml.OpenXmlElement) As System.String
        If element Is Nothing Then Return System.String.Empty
        Dim rgb As DocumentFormat.OpenXml.Drawing.RgbColorModelHex = element.Descendants(Of DocumentFormat.OpenXml.Drawing.RgbColorModelHex)().FirstOrDefault()
        If rgb?.Val IsNot Nothing Then Return "#" & rgb.Val.Value.ToUpperInvariant()
        Dim scheme As DocumentFormat.OpenXml.Drawing.SchemeColor = element.Descendants(Of DocumentFormat.OpenXml.Drawing.SchemeColor)().FirstOrDefault()
        If scheme?.Val IsNot Nothing Then Return "scheme:" & scheme.Val.Value.ToString()
        Return System.String.Empty
    End Function

    Private Function ExtractOutlineColor(ByVal shapeProperties As DocumentFormat.OpenXml.Presentation.ShapeProperties) As System.String
        Dim outline As DocumentFormat.OpenXml.Drawing.Outline = shapeProperties?.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.Outline)()
        Return ExtractSolidFillColor(outline)
    End Function

    Private Function ExtractBackgroundColor(ByVal background As DocumentFormat.OpenXml.Presentation.Background) As System.String
        If background Is Nothing Then Return System.String.Empty
        Return ExtractSolidFillColor(background)
    End Function

    Private Function ExtractEffectiveBackgroundColor(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As System.String
        If sp Is Nothing Then Return System.String.Empty
        Dim color As System.String = ExtractBackgroundColor(sp.Slide?.CommonSlideData?.Background)
        If Not System.String.IsNullOrWhiteSpace(color) Then Return color
        color = ExtractBackgroundColor(sp.SlideLayoutPart?.SlideLayout?.CommonSlideData?.Background)
        If Not System.String.IsNullOrWhiteSpace(color) Then Return color
        Return ExtractBackgroundColor(sp.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.Background)
    End Function

    Private Function ExtractTextStyleForJson(ByVal tb As DocumentFormat.OpenXml.Presentation.TextBody) As TextStyleJson
        Dim style As New TextStyleJson()
        If tb Is Nothing Then Return style

        Dim rp As DocumentFormat.OpenXml.Drawing.RunProperties = tb.Descendants(Of DocumentFormat.OpenXml.Drawing.RunProperties)().FirstOrDefault()
        If rp IsNot Nothing Then
            Dim latin As DocumentFormat.OpenXml.Drawing.LatinFont = rp.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.LatinFont)()
            If latin?.Typeface IsNot Nothing Then style.FontFamily = latin.Typeface.Value
            If rp.FontSize IsNot Nothing Then style.FontSize = rp.FontSize.Value / 100.0R
            If rp.Bold IsNot Nothing Then style.Bold = rp.Bold.Value
            If rp.Italic IsNot Nothing Then style.Italic = rp.Italic.Value
            style.Color = ExtractSolidFillColor(rp)
        End If

        Dim p As DocumentFormat.OpenXml.Drawing.Paragraph = tb.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().FirstOrDefault()
        Dim pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = p?.ParagraphProperties
        If pPr IsNot Nothing Then
            If pPr.Alignment IsNot Nothing Then style.Alignment = pPr.Alignment.Value.ToString()
            If pPr.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.NoBullet)() IsNot Nothing Then
                style.Bullet = False
            ElseIf pPr.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.CharacterBullet)() IsNot Nothing Then
                style.Bullet = True
                Dim cb As DocumentFormat.OpenXml.Drawing.CharacterBullet = pPr.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.CharacterBullet)()
                If cb?.Char IsNot Nothing Then style.BulletChar = cb.Char.Value
            End If
        End If
        Return style
    End Function

    Private Function MergeTextStyles(ByVal directStyle As TextStyleJson, ByVal inheritedStyle As TextStyleJson) As TextStyleJson
        If directStyle Is Nothing Then directStyle = New TextStyleJson()
        If inheritedStyle Is Nothing Then Return directStyle
        If System.String.IsNullOrWhiteSpace(directStyle.FontFamily) Then directStyle.FontFamily = inheritedStyle.FontFamily
        If Not directStyle.FontSize.HasValue Then directStyle.FontSize = inheritedStyle.FontSize
        If Not directStyle.Bold.HasValue Then directStyle.Bold = inheritedStyle.Bold
        If Not directStyle.Italic.HasValue Then directStyle.Italic = inheritedStyle.Italic
        If System.String.IsNullOrWhiteSpace(directStyle.Color) Then directStyle.Color = inheritedStyle.Color
        If System.String.IsNullOrWhiteSpace(directStyle.Alignment) Then directStyle.Alignment = inheritedStyle.Alignment
        If Not directStyle.Bullet.HasValue Then directStyle.Bullet = inheritedStyle.Bullet
        If System.String.IsNullOrWhiteSpace(directStyle.BulletChar) Then directStyle.BulletChar = inheritedStyle.BulletChar
        Return directStyle
    End Function

    ''' <summary>
    ''' Reads an OpenXML attribute by local name without requesting an
    ''' attribute/namespace combination that a strongly typed element rejects.
    ''' </summary>
    Private Function GetOpenXmlAttributeValueSafe(
        ByVal element As DocumentFormat.OpenXml.OpenXmlElement,
        ByVal localName As System.String) As System.String

        If element Is Nothing OrElse System.String.IsNullOrWhiteSpace(localName) Then
            Return System.String.Empty
        End If

        Try
            For Each attr As DocumentFormat.OpenXml.OpenXmlAttribute In element.GetAttributes()
                If System.String.Equals(attr.LocalName, localName, System.StringComparison.OrdinalIgnoreCase) Then
                    Return If(attr.Value, System.String.Empty)
                End If
            Next
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("Could not read OpenXML attribute '" & localName & "': " & ex.Message)
        End Try

        Return System.String.Empty
    End Function

    Private Function ExtractMasterTextStyleForJson(
        ByVal masterPart As DocumentFormat.OpenXml.Packaging.SlideMasterPart,
        ByVal role As System.String) As TextStyleJson

        Dim answer As New TextStyleJson()
        Dim textStyles As DocumentFormat.OpenXml.Presentation.TextStyles = masterPart?.SlideMaster?.TextStyles
        If textStyles Is Nothing Then Return answer

        Dim styleRoot As DocumentFormat.OpenXml.OpenXmlElement = Nothing
        If System.String.Equals(role, "title", System.StringComparison.OrdinalIgnoreCase) Then
            styleRoot = textStyles.TitleStyle
        ElseIf System.String.Equals(role, "body", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(role, "subtitle", System.StringComparison.OrdinalIgnoreCase) Then
            styleRoot = textStyles.BodyStyle
        Else
            styleRoot = textStyles.OtherStyle
        End If
        If styleRoot Is Nothing Then Return answer

        Dim level1 As DocumentFormat.OpenXml.OpenXmlElement = styleRoot.ChildElements.FirstOrDefault(Function(e) System.String.Equals(e.LocalName, "lvl1pPr", System.StringComparison.OrdinalIgnoreCase))
        If level1 Is Nothing Then level1 = styleRoot.ChildElements.FirstOrDefault()
        If level1 Is Nothing Then Return answer

        Dim defRp As DocumentFormat.OpenXml.Drawing.DefaultRunProperties = level1.Descendants(Of DocumentFormat.OpenXml.Drawing.DefaultRunProperties)().FirstOrDefault()
        If defRp IsNot Nothing Then
            If defRp.FontSize IsNot Nothing Then answer.FontSize = defRp.FontSize.Value / 100.0R
            If defRp.Bold IsNot Nothing Then answer.Bold = defRp.Bold.Value
            If defRp.Italic IsNot Nothing Then answer.Italic = defRp.Italic.Value
            Dim latin As DocumentFormat.OpenXml.Drawing.LatinFont = defRp.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.LatinFont)()
            If latin?.Typeface IsNot Nothing Then answer.FontFamily = latin.Typeface.Value
            answer.Color = ExtractSolidFillColor(defRp)
        End If

        ' DrawingML attributes such as algn/char are unqualified in the XML. Do not
        ' call GetAttribute with the DrawingML element namespace here: on strongly
        ' typed OpenXML elements the SDK can throw "The element does not allow the
        ' specified attribute." when the requested namespace is not legal for that
        ' attribute. Enumerating the attributes is version-safe and schema-safe.
        Dim alignmentValue As System.String = GetOpenXmlAttributeValueSafe(level1, "algn")
        If Not System.String.IsNullOrWhiteSpace(alignmentValue) Then answer.Alignment = alignmentValue
        If level1.ChildElements.Any(Function(e) e.LocalName = "buNone") Then answer.Bullet = False
        Dim buChar As DocumentFormat.OpenXml.OpenXmlElement = level1.ChildElements.FirstOrDefault(Function(e) e.LocalName = "buChar")
        If buChar IsNot Nothing Then
            answer.Bullet = True
            Dim bulletCharValue As System.String = GetOpenXmlAttributeValueSafe(buChar, "char")
            If Not System.String.IsNullOrWhiteSpace(bulletCharValue) Then answer.BulletChar = bulletCharValue
        End If
        Return answer
    End Function

    Private Function ExtractEffectiveTextStyleForShape(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal shp As DocumentFormat.OpenXml.Presentation.Shape) As TextStyleJson

        If shp Is Nothing Then Return New TextStyleJson()
        Dim role As System.String = "other"
        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        If ph IsNot Nothing Then role = ResolvePlaceholderRoleForJson(ph, If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty))

        Dim inherited As TextStyleJson = ExtractMasterTextStyleForJson(sp?.SlideLayoutPart?.SlideMasterPart, role)
        If sp?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree IsNot Nothing Then
            Dim layoutShape As DocumentFormat.OpenXml.Presentation.Shape = FindMatchingPlaceholderShape(sp.SlideLayoutPart.SlideLayout.CommonSlideData.ShapeTree, shp)
            If layoutShape IsNot Nothing Then inherited = MergeTextStyles(ExtractTextStyleForJson(layoutShape.TextBody), inherited)
        End If
        Return MergeTextStyles(ExtractTextStyleForJson(shp.TextBody), inherited)
    End Function

    Private Sub PopulateEffectivePlaceholderMetadata(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal phMeta As PlaceholderJson)

        If sp Is Nothing OrElse phMeta Is Nothing Then Return
        Dim shp As DocumentFormat.OpenXml.Presentation.Shape = FindShapeForPlaceholderMetadata(sp.Slide.CommonSlideData?.ShapeTree, phMeta)
        If shp Is Nothing Then Return
        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shp)
        ApplyEffectiveTransform(phMeta, xfrm)
        phMeta.TextStyle = ExtractEffectiveTextStyleForShape(sp, shp)
        ResolveTextStyleColorForJson(phMeta.TextStyle, sp?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme)
    End Sub

    Private Sub PopulateEffectivePlaceholderMetadata(
        ByVal lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
        ByVal phMeta As PlaceholderJson)

        If lp Is Nothing OrElse phMeta Is Nothing Then Return
        Dim shp As DocumentFormat.OpenXml.Presentation.Shape = FindShapeForPlaceholderMetadata(lp.SlideLayout?.CommonSlideData?.ShapeTree, phMeta)
        If shp Is Nothing Then Return
        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = shp.ShapeProperties?.Transform2D
        If xfrm Is Nothing AndAlso lp.SlideMasterPart IsNot Nothing Then
            Dim masterShape As DocumentFormat.OpenXml.Presentation.Shape = FindMatchingPlaceholderShape(lp.SlideMasterPart.SlideMaster?.CommonSlideData?.ShapeTree, shp)
            xfrm = masterShape?.ShapeProperties?.Transform2D
        End If
        ApplyEffectiveTransform(phMeta, xfrm)
        phMeta.TextStyle = MergeTextStyles(ExtractTextStyleForJson(shp.TextBody), ExtractMasterTextStyleForJson(lp.SlideMasterPart, phMeta.Role))
        ResolveTextStyleColorForJson(phMeta.TextStyle, lp?.SlideMasterPart?.ThemePart?.Theme)
    End Sub

    Private Sub ApplyEffectiveTransform(ByVal phMeta As PlaceholderJson, ByVal xfrm As DocumentFormat.OpenXml.Drawing.Transform2D)
        If phMeta Is Nothing OrElse xfrm Is Nothing Then Return
        If xfrm.Offset IsNot Nothing Then
            phMeta.EffectiveX = xfrm.Offset.X?.Value
            phMeta.EffectiveY = xfrm.Offset.Y?.Value
        End If
        If xfrm.Extents IsNot Nothing Then
            phMeta.EffectiveCx = xfrm.Extents.Cx?.Value
            phMeta.EffectiveCy = xfrm.Extents.Cy?.Value
        End If
    End Sub

    Private Function FindShapeForPlaceholderMetadata(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal phMeta As PlaceholderJson) As DocumentFormat.OpenXml.Presentation.Shape

        If tree Is Nothing OrElse phMeta Is Nothing Then Return Nothing
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            If phMeta.ShapeId.HasValue AndAlso shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id IsNot Nothing AndAlso
               shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value = phMeta.ShapeId.Value Then Return shp
        Next
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            If phMeta.Index.HasValue AndAlso ph.Index IsNot Nothing AndAlso ph.Index.Value = phMeta.Index.Value Then Return shp
            If Not System.String.IsNullOrWhiteSpace(phMeta.PlaceholderType) AndAlso ph.Type IsNot Nothing AndAlso
               System.String.Equals(ph.Type.Value.ToString(), phMeta.PlaceholderType, System.StringComparison.OrdinalIgnoreCase) Then Return shp
        Next
        Return Nothing
    End Function

    Private Function FindMatchingPlaceholderShape(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal sourceShape As DocumentFormat.OpenXml.Presentation.Shape) As DocumentFormat.OpenXml.Presentation.Shape

        If tree Is Nothing OrElse sourceShape Is Nothing Then Return Nothing
        Dim sourcePh As DocumentFormat.OpenXml.Presentation.PlaceholderShape = sourceShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        If sourcePh Is Nothing Then Return Nothing

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            If sourcePh.Index IsNot Nothing AndAlso ph.Index IsNot Nothing AndAlso sourcePh.Index.Value = ph.Index.Value Then Return shp
            If sourcePh.Type IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso sourcePh.Type.Value = ph.Type.Value Then Return shp
        Next
        Return Nothing
    End Function

    Private Function GetEffectiveTransformForShape(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal shp As DocumentFormat.OpenXml.Presentation.Shape) As DocumentFormat.OpenXml.Drawing.Transform2D

        If shp?.ShapeProperties?.Transform2D IsNot Nothing Then
            Return CType(shp.ShapeProperties.Transform2D.CloneNode(True), DocumentFormat.OpenXml.Drawing.Transform2D)
        End If
        Dim lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp?.SlideLayoutPart
        If lp IsNot Nothing Then
            Dim layoutShape As DocumentFormat.OpenXml.Presentation.Shape = FindMatchingPlaceholderShape(lp.SlideLayout?.CommonSlideData?.ShapeTree, shp)
            If layoutShape?.ShapeProperties?.Transform2D IsNot Nothing Then
                Return CType(layoutShape.ShapeProperties.Transform2D.CloneNode(True), DocumentFormat.OpenXml.Drawing.Transform2D)
            End If
            If lp.SlideMasterPart IsNot Nothing Then
                Dim masterShape As DocumentFormat.OpenXml.Presentation.Shape = FindMatchingPlaceholderShape(lp.SlideMasterPart.SlideMaster?.CommonSlideData?.ShapeTree, shp)
                If masterShape?.ShapeProperties?.Transform2D IsNot Nothing Then
                    Return CType(masterShape.ShapeProperties.Transform2D.CloneNode(True), DocumentFormat.OpenXml.Drawing.Transform2D)
                End If
            End If
        End If
        Return Nothing
    End Function

    Private Function BuildLayoutSignature(ByVal lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As System.String
        If lp?.SlideLayout?.CommonSlideData?.ShapeTree Is Nothing Then Return "empty"
        Dim counts As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
        For Each ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape In lp.SlideLayout.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)()
            Dim key As System.String = If(ph.Type IsNot Nothing, ph.Type.Value.ToString(), "implicit")
            If Not counts.ContainsKey(key) Then counts(key) = 0
            counts(key) += 1
        Next
        If counts.Count = 0 Then Return "blank"
        Return System.String.Join("+", counts.OrderBy(Function(kv) kv.Key).Select(Function(kv) kv.Key & If(kv.Value > 1, "x" & kv.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), System.String.Empty)))
    End Function

    Private Function ClassifyLayoutSemanticRole(ByVal lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As System.String
        If lp?.SlideLayout?.CommonSlideData?.ShapeTree Is Nothing Then Return "blank"

        Dim shapes As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
            lp.SlideLayout.CommonSlideData.ShapeTree.
                Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)().
                ToList()

        Dim hasTitle As System.Boolean = False
        Dim contentBodyCount As System.Int32 = 0
        Dim subtitleCount As System.Int32 = 0
        Dim heroImageCount As System.Int32 = 0
        Dim coverMetadataCount As System.Int32 = 0
        Dim closingCount As System.Int32 = 0
        Dim chartCount As System.Int32 = 0
        Dim tableCount As System.Int32 = 0

        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In shapes
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For

            Dim shapeName As System.String =
                If(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
            Dim semanticRole As System.String = ResolvePlaceholderSemanticRole(ph, shapeName)

            If System.String.Equals(semanticRole, "title", System.StringComparison.OrdinalIgnoreCase) Then hasTitle = True
            If System.String.Equals(semanticRole, "body", System.StringComparison.OrdinalIgnoreCase) Then contentBodyCount += 1
            If System.String.Equals(semanticRole, "subtitle", System.StringComparison.OrdinalIgnoreCase) Then subtitleCount += 1
            If System.String.Equals(semanticRole, "hero_image", System.StringComparison.OrdinalIgnoreCase) Then heroImageCount += 1
            If System.String.Equals(semanticRole, "presenter", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(semanticRole, "presenter_role", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(semanticRole, "location", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(semanticRole, "date", System.StringComparison.OrdinalIgnoreCase) Then
                coverMetadataCount += 1
            End If
            If System.String.Equals(semanticRole, "closing_message", System.StringComparison.OrdinalIgnoreCase) Then closingCount += 1

            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Chart Then chartCount += 1
            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Table Then tableCount += 1
        Next

        ' Picture placeholders are represented as p:pic rather than p:sp in many templates.
        For Each picture As DocumentFormat.OpenXml.Presentation.Picture In
            lp.SlideLayout.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Picture)()

            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                picture.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            Dim name As System.String = If(picture.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
            If System.String.Equals(ResolvePlaceholderSemanticRole(ph, name), "hero_image", System.StringComparison.OrdinalIgnoreCase) Then heroImageCount += 1
        Next

        If closingCount > 0 Then Return "closing"
        If hasTitle AndAlso subtitleCount > 0 AndAlso (heroImageCount > 0 OrElse coverMetadataCount >= 2) Then Return "cover"
        If hasTitle AndAlso contentBodyCount = 0 AndAlso heroImageCount = 0 AndAlso chartCount = 0 AndAlso tableCount = 0 Then Return "section"
        If chartCount > 0 OrElse tableCount > 0 Then Return "data"
        If heroImageCount > 0 AndAlso contentBodyCount <= 1 Then Return "visual"
        If contentBodyCount >= 2 Then Return "comparison_or_multi_content"
        If contentBodyCount = 1 Then Return "title_and_content"
        If Not hasTitle AndAlso contentBodyCount = 0 Then Return "blank"
        Return "mixed"
    End Function

    Private Function GetRecommendedUsesForLayout(ByVal semanticRole As System.String) As System.Collections.Generic.List(Of System.String)
        Select Case semanticRole
            Case "cover" : Return New System.Collections.Generic.List(Of System.String) From {"deck cover", "opening title"}
            Case "closing" : Return New System.Collections.Generic.List(Of System.String) From {"closing slide", "contact slide", "final call to action"}
            Case "section" : Return New System.Collections.Generic.List(Of System.String) From {"chapter divider", "section break", "single-message transition"}
            Case "data" : Return New System.Collections.Generic.List(Of System.String) From {"chart", "table", "data-heavy evidence"}
            Case "visual" : Return New System.Collections.Generic.List(Of System.String) From {"hero visual", "image-led explanation", "visual callout"}
            Case "comparison_or_multi_content" : Return New System.Collections.Generic.List(Of System.String) From {"comparison", "paired options", "two or more parallel concepts"}
            Case "title_and_content" : Return New System.Collections.Generic.List(Of System.String) From {"standard content", "key points", "single visual component in body region"}
            Case "blank" : Return New System.Collections.Generic.List(Of System.String) From {"custom visual composition", "diagram", "full-slide graphic"}
            Case Else : Return New System.Collections.Generic.List(Of System.String) From {"only when the placeholder structure exactly fits the intended content"}
        End Select
    End Function

    Private Function GetAvoidUsesForLayout(ByVal semanticRole As System.String) As System.Collections.Generic.List(Of System.String)
        Select Case semanticRole
            Case "cover" : Return New System.Collections.Generic.List(Of System.String) From {"normal content", "dense bullets", "data slides"}
            Case "closing" : Return New System.Collections.Generic.List(Of System.String) From {"normal analytical content", "dense comparison", "process diagram"}
            Case "section" : Return New System.Collections.Generic.List(Of System.String) From {"dense content", "multi-point explanation"}
            Case "data" : Return New System.Collections.Generic.List(Of System.String) From {"generic text-only slide when no data is present"}
            Case "comparison_or_multi_content" : Return New System.Collections.Generic.List(Of System.String) From {"single linear story", "content that does not have parallel parts"}
            Case Else : Return New System.Collections.Generic.List(Of System.String)()
        End Select
    End Function

    Private Function IsLayoutSafeForComponents(ByVal lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As System.Boolean
        If lp?.SlideLayout?.CommonSlideData?.ShapeTree Is Nothing Then Return True
        Dim role As System.String = ClassifyLayoutSemanticRole(lp)
        Return role = "blank" OrElse
               role = "title_and_content" OrElse
               role = "comparison_or_multi_content" OrElse
               role = "data"
    End Function

    Private Function ScoreLayoutQuality(ByVal lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart, ByVal usageCount As System.Int32) As System.Int32
        Dim score As System.Int32 = 35
        Dim role As System.String = ClassifyLayoutSemanticRole(lp)
        If role <> "mixed" Then score += 15
        If usageCount > 0 Then score += System.Math.Min(30, usageCount * 6)
        If lp?.SlideLayout?.CommonSlideData?.ShapeTree IsNot Nothing Then
            Dim placeholders As System.Int32 = lp.SlideLayout.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)().Count()
            If placeholders >= 1 AndAlso placeholders <= 4 Then score += 10
            Dim decorativeCount As System.Int32 = System.Linq.Enumerable.Count(
                lp.SlideLayout.CommonSlideData.ShapeTree.ChildElements.Cast(Of DocumentFormat.OpenXml.OpenXmlElement)(),
                Function(e) Not e.Descendants(Of DocumentFormat.OpenXml.Presentation.PlaceholderShape)().Any())
            If decorativeCount > 2 Then score += 10
        End If
        Return System.Math.Max(0, System.Math.Min(100, score))
    End Function

    Private Function ScoreSlideAsReusableTemplate(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal slideMeta As SlideJson,
        ByRef reason As System.String) As System.Int32

        Dim score As System.Int32 = 10
        Dim visuals As System.Collections.Generic.List(Of VisualElementJson) = If(slideMeta?.VisualElements, New System.Collections.Generic.List(Of VisualElementJson)())
        Dim slots As System.Collections.Generic.List(Of EditableSlotJson) = If(slideMeta?.EditableSlots, New System.Collections.Generic.List(Of EditableSlotJson)())
        Dim hasComplexGraphic As System.Boolean = visuals.Any(Function(v) v.Kind = "chart" OrElse v.Kind = "table" OrElse v.Kind = "diagram")
        Dim decorativeCount As System.Int32 = System.Linq.Enumerable.Count(visuals, Function(v) Not v.IsPlaceholder AndAlso (v.Kind = "shape" OrElse v.Kind = "picture"))

        If slots.Count >= 2 AndAlso slots.Count <= 8 Then score += 25
        If decorativeCount >= 2 Then score += 25
        If decorativeCount >= 5 Then score += 10
        If hasComplexGraphic Then score -= 35
        If slideMeta IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(slideMeta.Title) Then score += 5
        If visuals.Count <= 20 Then score += 10

        ' Many corporate template decks contain a gallery of intentionally empty sample
        ' slides. Those are valuable design exemplars even though they contain almost no
        ' substantive text and therefore no decorative slide-local shapes.
        Dim contentSlots As System.Collections.Generic.List(Of EditableSlotJson) =
            slots.Where(
                Function(slot)
                    Return slot IsNot Nothing AndAlso
                           Not System.String.Equals(slot.Role, "metadata", System.StringComparison.OrdinalIgnoreCase)
                End Function).
            ToList()

        Dim populatedContentSlots As System.Int32 =
            System.Linq.Enumerable.Count(
                contentSlots,
                Function(slot) Not System.String.IsNullOrWhiteSpace(slot.Text))

        If contentSlots.Count >= 2 AndAlso contentSlots.Count <= 12 AndAlso populatedContentSlots = 0 AndAlso Not hasComplexGraphic Then
            score += 30
        End If

        score = System.Math.Max(0, System.Math.Min(100, score))
        If score >= 70 Then
            reason = "Strong reusable visual pattern; clone only when the target content maps cleanly to the same slot structure"
        ElseIf score >= 45 Then
            reason = "Potential reusable pattern; use only for a close semantic and structural match"
        Else
            reason = "Prefer the slide layout rather than cloning this slide"
        End If
        Return score
    End Function

    Private Function BuildDesignProfile(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal result As PresentationJson) As DesignProfileJson

        Dim profile As New DesignProfileJson With {
            .Palette = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase),
            .ObservedColors = New System.Collections.Generic.List(Of System.String)(),
            .CommonShapeTypes = New System.Collections.Generic.List(Of System.String)(),
            .ExistingSlideCount = If(result?.Slides IsNot Nothing, result.Slides.Count, 0)
        }

        Dim master As DocumentFormat.OpenXml.Packaging.SlideMasterPart = PickMostUsedMaster(presPart)
        Dim theme As DocumentFormat.OpenXml.Drawing.Theme = master?.ThemePart?.Theme
        FillThemeProfile(theme, profile)
        profile.MasterBackgroundColor = ExtractBackgroundColor(master?.SlideMaster?.CommonSlideData?.Background)
        profile.MasterVisualElements = CollectVisualElementsFromShapeTree(master?.SlideMaster?.CommonSlideData?.ShapeTree)

        Dim colorCounts As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
        Dim shapeCounts As New System.Collections.Generic.Dictionary(Of System.String, System.Int32)(System.StringComparer.OrdinalIgnoreCase)
        Dim visualCount As System.Int32 = 0
        If result?.Slides IsNot Nothing Then
            For Each slideMeta As SlideJson In result.Slides
                If slideMeta?.VisualElements Is Nothing Then Continue For
                For Each v As VisualElementJson In slideMeta.VisualElements
                    visualCount += 1
                    If Not System.String.IsNullOrWhiteSpace(v.FillColor) AndAlso v.FillColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.FillColor)
                    If Not System.String.IsNullOrWhiteSpace(v.OutlineColor) AndAlso v.OutlineColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.OutlineColor)
                    If Not System.String.IsNullOrWhiteSpace(v.ShapeType) Then IncrementCount(shapeCounts, v.ShapeType)
                Next
            Next
        End If

        If result?.Layouts IsNot Nothing Then
            For Each layoutMeta As LayoutJson In result.Layouts
                If layoutMeta?.VisualElements Is Nothing Then Continue For
                For Each v As VisualElementJson In layoutMeta.VisualElements
                    If Not System.String.IsNullOrWhiteSpace(v.FillColor) AndAlso v.FillColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.FillColor)
                    If Not System.String.IsNullOrWhiteSpace(v.OutlineColor) AndAlso v.OutlineColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.OutlineColor)
                    If Not System.String.IsNullOrWhiteSpace(v.ShapeType) Then IncrementCount(shapeCounts, v.ShapeType)
                Next
            Next
        End If
        If profile.MasterVisualElements IsNot Nothing Then
            For Each v As VisualElementJson In profile.MasterVisualElements
                If Not System.String.IsNullOrWhiteSpace(v.FillColor) AndAlso v.FillColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.FillColor)
                If Not System.String.IsNullOrWhiteSpace(v.OutlineColor) AndAlso v.OutlineColor.StartsWith("#", System.StringComparison.Ordinal) Then IncrementCount(colorCounts, v.OutlineColor)
                If Not System.String.IsNullOrWhiteSpace(v.ShapeType) Then IncrementCount(shapeCounts, v.ShapeType)
            Next
        End If

        profile.ObservedColors = colorCounts.OrderByDescending(Function(kv) kv.Value).Take(8).Select(Function(kv) kv.Key).ToList()
        profile.CommonShapeTypes = shapeCounts.OrderByDescending(Function(kv) kv.Value).Take(8).Select(Function(kv) kv.Key).ToList()

        If profile.ExistingSlideCount = 0 Then
            profile.VisualDensity = "none"
        Else
            Dim perSlide As System.Double = visualCount / CDbl(System.Math.Max(1, profile.ExistingSlideCount))
            profile.VisualDensity = If(perSlide >= 8.0R, "high", If(perSlide >= 3.0R, "medium", "low"))
        End If

        Dim strength As System.Int32 = 15
        If profile.ExistingSlideCount > 0 Then strength += 25
        If result?.Layouts IsNot Nothing AndAlso result.Layouts.Count >= 4 Then strength += 10
        If visualCount >= 8 Then strength += 20
        If profile.ObservedColors.Count >= 2 Then strength += 10
        If Not System.String.IsNullOrWhiteSpace(profile.HeadingFont) AndAlso
           Not profile.HeadingFont.Equals("Aptos Display", System.StringComparison.OrdinalIgnoreCase) AndAlso
           Not profile.HeadingFont.Equals("Calibri Light", System.StringComparison.OrdinalIgnoreCase) Then strength += 10
        If result?.Layouts IsNot Nothing AndAlso result.Layouts.Any(Function(l) l.UsageCount >= 2) Then strength += 10
        profile.TemplateStrength = System.Math.Max(0, System.Math.Min(100, strength))

        Dim isGenericStarterDeck As System.Boolean =
            (profile.ExistingSlideCount = 0) OrElse
            (profile.ExistingSlideCount <= 1 AndAlso
             visualCount <= 2 AndAlso
             IsGenericOfficeTypography(profile))

        If isGenericStarterDeck Then
            profile.ModeHint = "blank_or_generic"
            profile.TemplateStrength = System.Math.Min(profile.TemplateStrength, 35)
            profile.Guidance = "Treat the stock Office starter deck as a blank canvas. Build a polished strategy-consulting visual system with restrained navy/blue/teal accents, deliberate slide backgrounds, message-led titles, charts, structured infographics, direct labels, and varied exhibits rather than default PowerPoint layouts or bullet-only slides"
        ElseIf profile.TemplateStrength >= 65 Then
            profile.ModeHint = "strong_existing_template"
            profile.Guidance = "Match the existing deck closely; prefer high-quality layouts and structurally matching reusable slides; use consulting-style charts and infographics inside safe content regions, always inheriting the deck typography and palette"
        Else
            profile.ModeHint = "light_template"
            profile.Guidance = "Respect the theme and recurring geometry, but improve weak layouts with restrained consulting-style charts, infographics, and stronger information hierarchy"
        End If

        AnalyzeSampleDeck(result, profile)
        Return profile
    End Function

    Private Sub AnalyzeSampleDeck(
        ByVal result As PresentationJson,
        ByVal profile As DesignProfileJson)

        If profile Is Nothing Then Return
        profile.SampleDeckLikelihood = "low"
        profile.SampleDeckConfidence = 0
        profile.SampleSlideKeys = New System.Collections.Generic.List(Of System.String)()
        profile.SampleDeckReason = System.String.Empty

        If result?.Slides Is Nothing OrElse result.Slides.Count < 4 Then Return

        Dim distinctLayouts As System.Int32 =
            result.Slides.
                Where(Function(slide) slide IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(slide.LayoutId)).
                Select(Function(slide) slide.LayoutId).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                Count()

        For Each slide As SlideJson In result.Slides
            If IsLikelySampleSlide(slide) AndAlso Not System.String.IsNullOrWhiteSpace(slide.SlideKey) Then
                profile.SampleSlideKeys.Add(slide.SlideKey)
            End If
        Next

        Dim ratio As System.Double = profile.SampleSlideKeys.Count / CDbl(System.Math.Max(1, result.Slides.Count))
        Dim diverseLibrary As System.Boolean = distinctLayouts >= 3 OrElse distinctLayouts >= CInt(System.Math.Ceiling(result.Slides.Count * 0.25R))

        If ratio >= 0.8R AndAlso diverseLibrary Then
            profile.SampleDeckLikelihood = "high"
            profile.SampleDeckConfidence = System.Math.Min(99, CInt(System.Math.Round(75.0R + ratio * 20.0R)))
            profile.SampleDeckReason = "Most pre-existing slides contain empty editable content slots and appear to form a diverse template/sample-slide library rather than narrative content"
        ElseIf ratio >= 0.6R AndAlso diverseLibrary Then
            profile.SampleDeckLikelihood = "medium"
            profile.SampleDeckConfidence = System.Math.Min(89, CInt(System.Math.Round(50.0R + ratio * 30.0R)))
            profile.SampleDeckReason = "A majority of pre-existing slides appear to be empty design examples; verify before deleting them"
        End If
    End Sub

    Private Function IsLikelySampleSlide(ByVal slide As SlideJson) As System.Boolean
        If slide Is Nothing OrElse slide.EditableSlots Is Nothing Then Return False

        Dim contentSlots As System.Collections.Generic.List(Of EditableSlotJson) =
            slide.EditableSlots.
                Where(
                    Function(slot)
                        If slot Is Nothing Then Return False
                        Dim sr As System.String = If(slot.SemanticRole, System.String.Empty)
                        If System.String.Equals(slot.Role, "metadata", System.StringComparison.OrdinalIgnoreCase) Then Return False
                        If System.String.Equals(sr, "footer", System.StringComparison.OrdinalIgnoreCase) OrElse
                           System.String.Equals(sr, "slide_number", System.StringComparison.OrdinalIgnoreCase) OrElse
                           System.String.Equals(sr, "date", System.StringComparison.OrdinalIgnoreCase) Then Return False
                        Return True
                    End Function).
                ToList()

        Dim populated As System.Int32 =
            System.Linq.Enumerable.Count(
                contentSlots,
                Function(slot) Not System.String.IsNullOrWhiteSpace(slot.Text))

        If populated > 0 Then Return False

        ' Empty slides using a reusable corporate layout are strong sample-library candidates.
        If slide.CloneAsTemplateScore >= 45 Then Return True
        Return contentSlots.Count >= 1
    End Function

    Private Function IsGenericOfficeTypography(ByVal profile As DesignProfileJson) As System.Boolean
        If profile Is Nothing Then Return True

        Dim heading As System.String = If(profile.HeadingFont, System.String.Empty).Trim()
        Dim body As System.String = If(profile.BodyFont, System.String.Empty).Trim()
        Dim genericFonts As System.String() = {
            "Aptos", "Aptos Display", "Calibri", "Calibri Light"
        }

        Dim headingGeneric As System.Boolean = System.String.IsNullOrWhiteSpace(heading) OrElse
            genericFonts.Any(Function(f) System.String.Equals(f, heading, System.StringComparison.OrdinalIgnoreCase))
        Dim bodyGeneric As System.Boolean = System.String.IsNullOrWhiteSpace(body) OrElse
            genericFonts.Any(Function(f) System.String.Equals(f, body, System.StringComparison.OrdinalIgnoreCase))

        Return headingGeneric AndAlso bodyGeneric
    End Function

    Private Sub IncrementCount(ByVal counts As System.Collections.Generic.Dictionary(Of System.String, System.Int32), ByVal key As System.String)
        If counts Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return
        If Not counts.ContainsKey(key) Then counts(key) = 0
        counts(key) += 1
    End Sub

    Private Function PickMostUsedMaster(ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart) As DocumentFormat.OpenXml.Packaging.SlideMasterPart
        If presPart Is Nothing Then Return Nothing
        Dim counts As New System.Collections.Generic.Dictionary(Of DocumentFormat.OpenXml.Packaging.SlideMasterPart, System.Int32)()
        If presPart.Presentation?.SlideIdList IsNot Nothing Then
            For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                Try
                    Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                    Dim sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart = sp?.SlideLayoutPart?.SlideMasterPart
                    If sm IsNot Nothing Then
                        If Not counts.ContainsKey(sm) Then counts(sm) = 0
                        counts(sm) += 1
                    End If
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Could not count slide master usage: " & ex.Message)
                End Try
            Next
        End If
        If counts.Count > 0 Then Return counts.OrderByDescending(Function(kv) kv.Value).First().Key
        Return presPart.SlideMasterParts.FirstOrDefault()
    End Function

    Private Sub FillThemeProfile(ByVal theme As DocumentFormat.OpenXml.Drawing.Theme, ByVal profile As DesignProfileJson)
        If profile Is Nothing Then Return
        If theme?.ThemeElements?.ColorScheme IsNot Nothing Then
            Dim cs As DocumentFormat.OpenXml.Drawing.ColorScheme = theme.ThemeElements.ColorScheme
            profile.Palette("dark1") = ExtractThemeColor(cs.Dark1Color, "#0F172A")
            profile.Palette("light1") = ExtractThemeColor(cs.Light1Color, "#F8FAFC")
            profile.Palette("dark2") = ExtractThemeColor(cs.Dark2Color, "#334155")
            profile.Palette("light2") = ExtractThemeColor(cs.Light2Color, "#E2E8F0")
            profile.Palette("accent1") = ExtractThemeColor(cs.Accent1Color, "#2563EB")
            profile.Palette("accent2") = ExtractThemeColor(cs.Accent2Color, "#0F766E")
            profile.Palette("accent3") = ExtractThemeColor(cs.Accent3Color, "#D97706")
            profile.Palette("accent4") = ExtractThemeColor(cs.Accent4Color, "#7C3AED")
            profile.Palette("accent5") = ExtractThemeColor(cs.Accent5Color, "#DB2777")
            profile.Palette("accent6") = ExtractThemeColor(cs.Accent6Color, "#0891B2")
        Else
            profile.Palette("dark1") = "#0F172A"
            profile.Palette("light1") = "#F8FAFC"
            profile.Palette("accent1") = "#2563EB"
            profile.Palette("accent2") = "#0F766E"
            profile.Palette("accent3") = "#D97706"
        End If

        If theme?.ThemeElements?.FontScheme IsNot Nothing Then
            Dim fs As DocumentFormat.OpenXml.Drawing.FontScheme = theme.ThemeElements.FontScheme
            profile.HeadingFont = If(fs.MajorFont?.LatinFont?.Typeface?.Value, System.String.Empty)
            profile.BodyFont = If(fs.MinorFont?.LatinFont?.Typeface?.Value, System.String.Empty)
        End If
        If System.String.IsNullOrWhiteSpace(profile.HeadingFont) Then profile.HeadingFont = "Aptos Display"
        If System.String.IsNullOrWhiteSpace(profile.BodyFont) Then profile.BodyFont = "Aptos"
    End Sub

    Private Function ExtractThemeColor(ByVal element As DocumentFormat.OpenXml.OpenXmlElement, ByVal fallback As System.String) As System.String
        If element Is Nothing Then Return fallback
        Dim rgb As DocumentFormat.OpenXml.Drawing.RgbColorModelHex = element.Descendants(Of DocumentFormat.OpenXml.Drawing.RgbColorModelHex)().FirstOrDefault()
        If rgb?.Val IsNot Nothing Then Return "#" & rgb.Val.Value.ToUpperInvariant()
        Dim sys As DocumentFormat.OpenXml.Drawing.SystemColor = element.Descendants(Of DocumentFormat.OpenXml.Drawing.SystemColor)().FirstOrDefault()
        If sys?.LastColor IsNot Nothing Then Return "#" & sys.LastColor.Value.ToUpperInvariant()
        Return fallback
    End Function

#End Region

    ''' <summary>
    ''' Analyzes a layout part to determine which placeholder types are present.
    ''' </summary>
    ''' <param name="lp">The slide layout part to analyze.</param>
    ''' <returns>LayoutInfo with flags for title, subtitle, and body placeholders.</returns>
    Private Function AnalyzeLayoutPlaceholders(lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As LayoutInfo
        Dim li As New LayoutInfo()
        Dim tree = lp?.SlideLayout?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return li

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph = shp?.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            If ph.Type Is Nothing Then
                ' in PPT the "subtitle" box is frequently a placeholder with no explicit type but index=1
                If ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then
                    li.HasSubTitle = True
                End If
                Continue For
            End If
            Select Case ph.Type.Value
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title
                    li.HasTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle
                    li.HasCenteredTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                    li.HasSubTitle = True
                Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body
                    li.HasBody = True
            End Select
        Next

        Return li
    End Function

    ''' <summary>
    ''' Extracts the title text from a slide's title placeholder.
    ''' </summary>
    ''' <param name="sp">The slide part to search.</param>
    ''' <returns>Title text or empty string if not found.</returns>
    Private Function GetSlideTitle(sp As SlidePart) As String
        If sp Is Nothing OrElse sp.Slide Is Nothing OrElse
       sp.Slide.CommonSlideData Is Nothing OrElse
       sp.Slide.CommonSlideData.ShapeTree Is Nothing Then
            Return String.Empty
        End If

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In
        sp.Slide.CommonSlideData.ShapeTree.ChildElements _
          .OfType(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim nv = shp.NonVisualShapeProperties
            If nv IsNot Nothing AndAlso nv.ApplicationNonVisualDrawingProperties IsNot Nothing Then
                Dim ph = nv.ApplicationNonVisualDrawingProperties.PlaceholderShape
                If ph IsNot Nothing AndAlso
               (ph.Type Is Nothing OrElse
                ph.Type.Value = PlaceholderValues.Title OrElse
                ph.Type.Value = PlaceholderValues.CenteredTitle) Then
                    Return If(shp.TextBody IsNot Nothing, shp.TextBody.InnerText, String.Empty)
                End If
            End If
        Next
        Return String.Empty
    End Function


    ''' <summary>
    ''' Gets the human-readable name of a slide layout.
    ''' </summary>
    ''' <param name="layoutPart">The layout part.</param>
    ''' <returns>Layout name or URI string as fallback.</returns>
    Private Function GetLayoutName(layoutPart As SlideLayoutPart) As String
        If layoutPart Is Nothing Then Return String.Empty
        If layoutPart.SlideLayout IsNot Nothing AndAlso
       layoutPart.SlideLayout.CommonSlideData IsNot Nothing Then
            Dim nm = layoutPart.SlideLayout.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then Return nm
        End If
        Return If(layoutPart.Uri IsNot Nothing, layoutPart.Uri.ToString(), String.Empty)
    End Function

    ''' <summary>
    ''' Gets the human-readable name of a slide master.
    ''' </summary>
    ''' <param name="smPart">The slide master part.</param>
    ''' <returns>Master name or URI string as fallback.</returns>
    Private Function GetMasterName(smPart As SlideMasterPart) As String
        If smPart Is Nothing Then Return String.Empty
        If smPart.SlideMaster IsNot Nothing AndAlso
       smPart.SlideMaster.CommonSlideData IsNot Nothing Then
            Dim nm = smPart.SlideMaster.CommonSlideData.Name
            If Not String.IsNullOrWhiteSpace(nm) Then Return nm
        End If
        Return If(smPart.Uri IsNot Nothing, smPart.Uri.ToString(), String.Empty)
    End Function

    ''' <summary>
    ''' Sanitizes a string for use as a slide key by replacing non-alphanumeric characters.
    ''' </summary>
    ''' <param name="s">The string to sanitize.</param>
    ''' <returns>Sanitized string with only letters, digits, and hyphens.</returns>
    Private Function SanitizeKey(s As String) As String
        Return New String(
            s.Select(Function(ch) If(Char.IsLetterOrDigit(ch), ch, "-"c)).ToArray()
        )
    End Function

    ''' <summary>
    ''' Cleans a raw string to extract valid JSON by finding matching braces/brackets.
    ''' </summary>
    ''' <param name="raw">The raw string potentially containing JSON.</param>
    ''' <returns>Cleaned JSON string or trimmed original.</returns>
    Public Function CleanJsonString(raw As String) As String
        If String.IsNullOrEmpty(raw) Then
            Return String.Empty
        End If

        ' Look for object vs. array start
        Dim firstObj = raw.IndexOf("{"c)
        Dim firstArr = raw.IndexOf("["c)
        Dim startIdx As Integer
        Dim openChar As Char
        Dim closeChar As Char

        If firstObj >= 0 AndAlso (firstObj < firstArr OrElse firstArr = -1) Then
            startIdx = firstObj
            openChar = "{"c
            closeChar = "}"c
        ElseIf firstArr >= 0 Then
            startIdx = firstArr
            openChar = "["c
            closeChar = "]"c
        Else
            ' No JSON delimiters found – just return trimmed
            Return raw.Trim()
        End If

        ' Find the last matching closing brace/bracket
        Dim lastIdx = raw.LastIndexOf(closeChar)
        If lastIdx > startIdx Then
            Return raw.Substring(startIdx, lastIdx - startIdx + 1).Trim()
        Else
            ' Malformed or unmatched – fallback to trimming
            Return raw.Trim()
        End If
    End Function


    ''' <summary>
    ''' Applies an AI-generated action plan to modify a PowerPoint presentation.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file to modify.</param>
    ''' <param name="planJson">JSON string containing the action plan with slide operations.</param>
    ''' <returns>True if all actions applied successfully; False on error.</returns>
    ''' <summary>
    ''' Applies an AI-generated action plan to modify a PowerPoint presentation.
    ''' Version 2 keeps the version-1 element schema compatible and adds robust layout resolution,
    ''' correct before/after anchors, reusable-slide cloning, built-in icons, and high-level visual components.
    ''' </summary>
    Public Function ApplyPlanToPresentation(ByVal pptxPath As System.String, ByVal planJson As System.String) As System.Boolean
        Try
            If Not System.IO.File.Exists(pptxPath) Then
                ShowCustomMessageBox($"Your file '{pptxPath}' was no longer found - aborting.")
                Return False
            End If

            Dim errorMessages As New System.Collections.Generic.List(Of System.String)()
            Using planDoc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse(planJson)
                Dim root As System.Text.Json.JsonElement = planDoc.RootElement
                Dim actionsEl As System.Text.Json.JsonElement
                If Not root.TryGetProperty("actions", actionsEl) OrElse actionsEl.ValueKind <> System.Text.Json.JsonValueKind.Array Then
                    ShowCustomMessageBox("An internal error occurred when amending your slidedeck (the AI sent instructions missing the required 'actions' array).")
                    Return False
                End If

                Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
                    DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, True)

                    Dim presPart As DocumentFormat.OpenXml.Packaging.PresentationPart = presDoc.PresentationPart
                    If presPart Is Nothing Then
                        ShowCustomMessageBox("A presentation is missing in the file you have provided; you may have to include at least one slide.")
                        Return False
                    End If

                    EnsureSlideIdList(presPart)
                    Dim idx As DeckIndex = BuildDeckIndex(presPart)
                    Dim originalSlideIds As New System.Collections.Generic.HashSet(Of UInteger)(idx.SlideKeyById.Values)
                    Dim originalSlideKeys As New System.Collections.Generic.HashSet(Of System.String)(idx.SlideKeyById.Keys, System.StringComparer.OrdinalIgnoreCase)
                    Dim runtimeSampleKeys As System.Collections.Generic.HashSet(Of System.String) = DetectRuntimeSampleSlideKeys(presPart)
                    Dim lastInsertedKey As System.String = System.String.Empty
                    Dim runtimeDesignMode As System.String = DetermineRuntimeDesignMode(presPart)

                    For Each actElem As System.Text.Json.JsonElement In actionsEl.EnumerateArray()
                        Dim opEl As System.Text.Json.JsonElement
                        If Not actElem.TryGetProperty("op", opEl) OrElse opEl.ValueKind <> System.Text.Json.JsonValueKind.String Then Continue For
                        If Not System.String.Equals(opEl.GetString(), "add_slide", System.StringComparison.OrdinalIgnoreCase) Then Continue For

                        Try
                            Dim anchorMode As System.String = "at_end"
                            Dim anchorKey As System.String = System.String.Empty
                            Dim anchorEl As System.Text.Json.JsonElement
                            If actElem.TryGetProperty("anchor", anchorEl) AndAlso anchorEl.ValueKind = System.Text.Json.JsonValueKind.Object Then
                                Dim modeEl As System.Text.Json.JsonElement
                                If anchorEl.TryGetProperty("mode", modeEl) AndAlso modeEl.ValueKind = System.Text.Json.JsonValueKind.String Then anchorMode = modeEl.GetString()
                                Dim byEl As System.Text.Json.JsonElement
                                If anchorEl.TryGetProperty("by", byEl) AndAlso byEl.ValueKind = System.Text.Json.JsonValueKind.Object Then
                                    Dim keyEl As System.Text.Json.JsonElement
                                    If byEl.TryGetProperty("slideKey", keyEl) AndAlso keyEl.ValueKind = System.Text.Json.JsonValueKind.String Then anchorKey = keyEl.GetString()
                                End If
                            End If

                            If System.String.Equals(anchorKey, "lastInserted", System.StringComparison.OrdinalIgnoreCase) Then anchorKey = lastInsertedKey
                            Dim anchorId As UInteger = 0UI
                            If Not System.String.IsNullOrWhiteSpace(anchorKey) AndAlso idx.SlideKeyById.ContainsKey(anchorKey) Then anchorId = idx.SlideKeyById(anchorKey)

                            Dim newSp As DocumentFormat.OpenXml.Packaging.SlidePart = Nothing
                            Dim sourceKeyEl As System.Text.Json.JsonElement
                            If actElem.TryGetProperty("sourceSlideKey", sourceKeyEl) AndAlso sourceKeyEl.ValueKind = System.Text.Json.JsonValueKind.String AndAlso Not System.String.IsNullOrWhiteSpace(sourceKeyEl.GetString()) Then
                                Dim sourceSp As DocumentFormat.OpenXml.Packaging.SlidePart = FindSlidePartByKey(presPart, sourceKeyEl.GetString())
                                If sourceSp Is Nothing Then Throw New System.Exception("sourceSlideKey does not identify an existing slide.")
                                newSp = CloneExistingSlideAsTemplate(presPart, sourceSp)
                                ClearShapeTextByIds(newSp, actElem)
                            Else
                                Dim targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = ResolveLayoutFromAction(presPart, actElem)

                                ' Covers and closing slides frequently carry slide-local photography or other
                                ' specimen artwork in corporate template galleries. Prefer cloning the empty
                                ' sample slide for these special layouts so such visuals cannot disappear merely
                                ' because they are not encoded directly on the SlideLayoutPart.
                                Dim sampleSource As DocumentFormat.OpenXml.Packaging.SlidePart = Nothing
                                If ShouldPreferSampleSlideClone(targetLayout) Then
                                    sampleSource = FindRuntimeSampleSlideForLayout(presPart, targetLayout, runtimeSampleKeys)
                                End If

                                If sampleSource IsNot Nothing Then
                                    newSp = CloneExistingSlideAsTemplate(presPart, sampleSource)
                                    ClearAllTemplateCloneEditableText(newSp)
                                Else
                                    newSp = CloneTemplateSlide(presPart, targetLayout)
                                End If
                            End If

                            Dim slideStyle As System.String = JsonString(actElem, "slideStyle", System.String.Empty)
                            If System.String.Equals(runtimeDesignMode, "blank_or_generic", System.StringComparison.OrdinalIgnoreCase) Then
                                If System.String.IsNullOrWhiteSpace(slideStyle) Then slideStyle = InferConsultingSlideStyle(newSp)
                                ApplyConsultingSlideStyle(presPart, newSp, slideStyle)
                            End If

                            Dim newId As UInteger = InsertAtAnchor(presPart, anchorMode, anchorId, newSp)

                            Dim elementsEl As System.Text.Json.JsonElement
                            If actElem.TryGetProperty("elements", elementsEl) AndAlso elementsEl.ValueKind = System.Text.Json.JsonValueKind.Array Then
                                For Each el As System.Text.Json.JsonElement In elementsEl.EnumerateArray()
                                    Dim typeEl As System.Text.Json.JsonElement
                                    If Not el.TryGetProperty("type", typeEl) OrElse typeEl.ValueKind <> System.Text.Json.JsonValueKind.String Then Continue For
                                    Dim elementType As System.String = typeEl.GetString().ToLowerInvariant()

                                    Select Case elementType
                                        Case "title"
                                            Dim textEl As System.Text.Json.JsonElement
                                            If el.TryGetProperty("text", textEl) Then SetTitle(newSp, If(textEl.GetString(), System.String.Empty), el)
                                        Case "shape"
                                            AddShape(presPart, newSp, el)
                                        Case "svg_icon", "svg_graphic"
                                            AddSvgIcon(presPart, newSp, el)
                                        Case "icon"
                                            AddBuiltinIcon(presPart, newSp, el)
                                        Case "component"
                                            AddComponent(presPart, newSp, el)
                                        Case "text"
                                            If el.TryGetProperty("transform", Nothing) Then
                                                CreateFreestandingTextBox(presPart, newSp, el)
                                            Else
                                                Dim textEl As System.Text.Json.JsonElement
                                                If el.TryGetProperty("text", textEl) Then
                                                    Dim placeholderEl As System.Text.Json.JsonElement
                                                    If el.TryGetProperty("placeholder", placeholderEl) Then
                                                        SetTextWithPlaceholder(newSp, placeholderEl, If(textEl.GetString(), System.String.Empty), el)
                                                    End If
                                                End If
                                            End If
                                        Case "bullet_text"
                                            If el.TryGetProperty("transform", Nothing) Then
                                                CreateFreestandingTextBox(presPart, newSp, el)
                                            Else
                                                SetBulletsWithPlaceholder(newSp, el)
                                            End If
                                    End Select
                                Next
                            End If

                            If System.String.Equals(runtimeDesignMode, "blank_or_generic", System.StringComparison.OrdinalIgnoreCase) Then
                                ApplyConsultingPlaceholderTypography(newSp, slideStyle)
                            End If

                            EnsureGeneratedCoverMetadata(newSp)

                            Dim notesEl As System.Text.Json.JsonElement
                            If actElem.TryGetProperty("notes", notesEl) AndAlso notesEl.ValueKind = System.Text.Json.JsonValueKind.String Then SetSpeakerNotes(newSp, notesEl.GetString())

                            newSp.Slide.Save()
                            presPart.Presentation.Save()
                            idx = BuildDeckIndex(presPart)
                            lastInsertedKey = GetSlideKey(newSp, newId)

                        Catch ex As System.Exception
                            System.Diagnostics.Debug.WriteLine("Error creating slide: " & ex.Message)
                            errorMessages.Add("Could not implement one slide instruction: " & ex.Message)
                        End Try
                    Next

                    HandleSampleSlideCleanupRequest(
                        presPart,
                        root,
                        originalSlideIds,
                        originalSlideKeys,
                        runtimeSampleKeys)

                    For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                        Dim spPart As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                        If spPart IsNot Nothing AndAlso spPart.NotesSlidePart Is Nothing Then SetSpeakerNotes(spPart, System.String.Empty)
                    Next
                    presPart.Presentation.Save()
                End Using
            End Using

            If errorMessages.Count > 0 Then
                ShowCustomMessageBox("Several errors occurred while applying the AI's slide instructions (the deck may still have been modified partially):" & vbCrLf & vbCrLf & System.String.Join(vbCrLf, errorMessages))
                Return False
            End If
            Return True

        Catch ex As System.Text.Json.JsonException
            ShowCustomMessageBox("The AI has sent an invalid instruction on how to build the slides: " & ex.Message)
            Return False
        Catch ex As DocumentFormat.OpenXml.Packaging.OpenXmlPackageException
            ShowCustomMessageBox("A PowerPoint file error occurred: " & ex.Message)
            Return False
        Catch ex As System.Exception
            ShowCustomMessageBox("An unexpected error occurred when amending your slidedeck: " & ex.Message)
            Return False
        End Try
    End Function

    Private Function DetectRuntimeSampleSlideKeys(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart) As System.Collections.Generic.HashSet(Of System.String)

        Dim result As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
        If presPart?.Presentation?.SlideIdList Is Nothing Then Return result

        For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
            If sid.RelationshipId Is Nothing Then Continue For
            Try
                Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                If sp Is Nothing Then Continue For
                If IsRuntimeLikelySampleSlide(sp) Then result.Add(GetSlideKey(sp, sid.Id.Value))
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Could not inspect sample-slide candidate: " & ex.Message)
            End Try
        Next

        Return result
    End Function

    Private Function IsRuntimeLikelySampleSlide(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As System.Boolean

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return False

        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For

            Dim semanticRole As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shape)
            If System.String.Equals(semanticRole, "footer", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(semanticRole, "slide_number", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(semanticRole, "date", System.StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            Dim text As System.String = If(shape.TextBody IsNot Nothing, ExtractTextFromTextContainer(shape.TextBody), System.String.Empty)
            If Not System.String.IsNullOrWhiteSpace(text) Then Return False
        Next

        Return True
    End Function

    Private Sub HandleSampleSlideCleanupRequest(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal root As System.Text.Json.JsonElement,
        ByVal originalSlideIds As System.Collections.Generic.HashSet(Of UInteger),
        ByVal originalSlideKeys As System.Collections.Generic.HashSet(Of System.String),
        ByVal runtimeSampleKeys As System.Collections.Generic.HashSet(Of System.String))

        If presPart Is Nothing OrElse root.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return

        Dim cleanupEl As System.Text.Json.JsonElement
        If Not root.TryGetProperty("cleanup", cleanupEl) OrElse cleanupEl.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return

        Dim offer As System.Boolean = False
        Dim tmp As System.Text.Json.JsonElement
        If cleanupEl.TryGetProperty("offerDeleteSampleSlides", tmp) AndAlso
           (tmp.ValueKind = System.Text.Json.JsonValueKind.True OrElse tmp.ValueKind = System.Text.Json.JsonValueKind.False) Then
            offer = tmp.GetBoolean()
        End If
        If Not offer Then Return

        Dim requestedKeys As New System.Collections.Generic.List(Of System.String)()
        If cleanupEl.TryGetProperty("sampleSlideKeys", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Array Then
            For Each item As System.Text.Json.JsonElement In tmp.EnumerateArray()
                If item.ValueKind = System.Text.Json.JsonValueKind.String AndAlso Not System.String.IsNullOrWhiteSpace(item.GetString()) Then
                    requestedKeys.Add(item.GetString())
                End If
            Next
        End If
        If requestedKeys.Count = 0 Then Return

        Dim safeKeys As System.Collections.Generic.List(Of System.String) =
            requestedKeys.
                Where(
                    Function(key)
                        Return originalSlideKeys.Contains(key) AndAlso runtimeSampleKeys.Contains(key)
                    End Function).
                Distinct(System.StringComparer.OrdinalIgnoreCase).
                ToList()

        ' The model may omit the first/title specimen because it contains a date field and
        ' therefore looks superficially non-empty. When it has correctly identified a large
        ' sample-slide gallery, augment its list with every original slide that the renderer
        ' independently confirmed to be empty/sample-safe. This remains conservative because
        ' IsRuntimeLikelySampleSlide rejects any slide containing real title/body content.
        Dim broadSampleCleanupRequest As System.Boolean =
            requestedKeys.Count >= System.Math.Max(3, CInt(System.Math.Floor(originalSlideKeys.Count * 0.5R)))

        If broadSampleCleanupRequest Then
            For Each runtimeKey As System.String In runtimeSampleKeys
                If originalSlideKeys.Contains(runtimeKey) AndAlso
                   Not System.Linq.Enumerable.Contains(safeKeys, runtimeKey, System.StringComparer.OrdinalIgnoreCase) Then
                    safeKeys.Add(runtimeKey)
                End If
            Next
        End If

        If safeKeys.Count = 0 Then Return

        Dim question As System.String =
            "The original presentation appears to contain " &
            safeKeys.Count.ToString(System.Globalization.CultureInfo.InvariantCulture) &
            " empty sample/template slide" &
            If(safeKeys.Count = 1, "", "s") &
            ". Delete those original sample slides now and keep the newly generated presentation?"

        Dim answer As System.Int32 = ShowCustomYesNoBox(question, "Yes", "No")
        Select Case answer
            Case 1
                DeleteOriginalSlidesByKeys(presPart, safeKeys, originalSlideIds)
            Case 2
                ' Keep the original sample slides.
            Case Else
                ' 0 = Abort/closed dialog. Generation has already completed, so abort cleanup
                ' only and keep both the new slides and the original sample slides.
        End Select
    End Sub

    Private Sub DeleteOriginalSlidesByKeys(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal slideKeys As System.Collections.Generic.IEnumerable(Of System.String),
        ByVal originalSlideIds As System.Collections.Generic.HashSet(Of UInteger))

        If presPart?.Presentation?.SlideIdList Is Nothing OrElse slideKeys Is Nothing Then Return
        Dim index As DeckIndex = BuildDeckIndex(presPart)
        Dim idsToDelete As New System.Collections.Generic.HashSet(Of UInteger)()

        For Each key As System.String In slideKeys
            If System.String.IsNullOrWhiteSpace(key) OrElse Not index.SlideKeyById.ContainsKey(key) Then Continue For
            Dim id As UInteger = index.SlideKeyById(key)
            If originalSlideIds.Contains(id) Then idsToDelete.Add(id)
        Next

        Dim slideList As DocumentFormat.OpenXml.Presentation.SlideIdList = presPart.Presentation.SlideIdList
        If System.Linq.Enumerable.Count(slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()) - idsToDelete.Count < 1 Then Return

        For Each slideId As UInteger In idsToDelete
            Dim sid As DocumentFormat.OpenXml.Presentation.SlideId =
                slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)().
                    FirstOrDefault(Function(item) item.Id IsNot Nothing AndAlso item.Id.Value = slideId)
            If sid Is Nothing Then Continue For

            Dim slidePart As DocumentFormat.OpenXml.Packaging.SlidePart = Nothing
            If sid.RelationshipId IsNot Nothing Then
                slidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
            End If

            sid.Remove()
            If slidePart IsNot Nothing Then
                Try
                    presPart.DeletePart(slidePart)
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Could not remove sample slide part: " & ex.Message)
                End Try
            End If
        Next

        presPart.Presentation.Save()
    End Sub

    Private Sub EnsureGeneratedCoverMetadata(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart)

        If sp?.SlideLayoutPart Is Nothing Then Return
        If Not System.String.Equals(ClassifyLayoutSemanticRole(sp.SlideLayoutPart), "cover", System.StringComparison.OrdinalIgnoreCase) Then Return

        Dim dateShape As DocumentFormat.OpenXml.Presentation.Shape = FindShapeBySemanticRole(sp, "date")
        If dateShape Is Nothing Then Return

        ' A cloned sample cover may contain the date on which the template was authored.
        ' A newly generated cover should always carry the current presentation date unless
        ' the caller later explicitly overwrites the date placeholder.
        Dim dateText As System.String = System.DateTime.Now.ToString("d. MMMM yyyy", System.Globalization.CultureInfo.CurrentCulture)
        Using emptyDoc As System.Text.Json.JsonDocument = System.Text.Json.JsonDocument.Parse("{}")
            SetShapeSingleTextPreserveStyle(dateShape, dateText, emptyDoc.RootElement, forceNoBullet:=True)
            ApplyGeneratedPlaceholderTextFit(sp, dateShape, dateText)
        End Using
        sp.Slide.Save()
    End Sub

    Private Function FindShapeBySemanticRole(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal semanticRole As System.String) As DocumentFormat.OpenXml.Presentation.Shape

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return Nothing
        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            Dim role As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shape)
            If System.String.Equals(role, semanticRole, System.StringComparison.OrdinalIgnoreCase) Then Return shape
        Next
        Return Nothing
    End Function

    Private Function IsShapeTextEmpty(ByVal shape As DocumentFormat.OpenXml.Presentation.Shape) As System.Boolean
        If shape?.TextBody Is Nothing Then Return True
        Return Not shape.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().
            Any(Function(t) Not System.String.IsNullOrWhiteSpace(t.Text))
    End Function

    Private Function DetermineRuntimeDesignMode(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart) As System.String

        If presPart Is Nothing OrElse presPart.Presentation Is Nothing Then Return "blank_or_generic"

        Dim slideCount As System.Int32 = 0
        If presPart.Presentation.SlideIdList IsNot Nothing Then
            slideCount = presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)().Count()
        End If
        If slideCount = 0 Then Return "blank_or_generic"

        Dim customVisualCount As System.Int32 = 0
        If presPart.Presentation.SlideIdList IsNot Nothing Then
            For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                Try
                    Dim slidePart As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                    Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = slidePart?.Slide?.CommonSlideData?.ShapeTree
                    If tree Is Nothing Then Continue For

                    For Each child As DocumentFormat.OpenXml.OpenXmlElement In tree.ChildElements
                        If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                            Dim shape As DocumentFormat.OpenXml.Presentation.Shape = CType(child, DocumentFormat.OpenXml.Presentation.Shape)
                            Dim placeholder As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                            If placeholder Is Nothing Then customVisualCount += 1
                        ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.Picture OrElse
                               TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame OrElse
                               TypeOf child Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                            customVisualCount += 1
                        End If
                    Next
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine("Could not inspect starter slide visuals: " & ex.Message)
                End Try
            Next
        End If

        Dim profile As New DesignProfileJson With {
            .Palette = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        }
        Dim master As DocumentFormat.OpenXml.Packaging.SlideMasterPart = PickMostUsedMaster(presPart)
        FillThemeProfile(master?.ThemePart?.Theme, profile)

        If slideCount <= 1 AndAlso customVisualCount <= 2 AndAlso IsGenericOfficeTypography(profile) Then
            Return "blank_or_generic"
        End If

        Return "existing_template"
    End Function

    Private Function InferConsultingSlideStyle(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As System.String

        Dim layout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = sp?.SlideLayoutPart
        If layout Is Nothing Then Return "content_light"

        Dim layoutName As System.String = GetLayoutName(layout).ToLowerInvariant()
        If layoutName.Contains("section") OrElse layoutName.Contains("chapter") Then Return "section_dark"

        Dim info As LayoutInfo = AnalyzeLayoutPlaceholders(layout)
        If (info.HasTitle OrElse info.HasCenteredTitle) AndAlso info.HasSubTitle AndAlso Not info.HasBody Then
            Return "cover_dark"
        End If

        Return "content_light"
    End Function

    Private Sub ApplyConsultingSlideStyle(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal slideStyle As System.String)

        If presPart Is Nothing OrElse sp Is Nothing Then Return
        Dim styleName As System.String = If(slideStyle, System.String.Empty).Trim().ToLowerInvariant()
        If System.String.IsNullOrWhiteSpace(styleName) Then styleName = "content_light"
        If styleName = "native" Then Return

        Dim background As System.String
        Select Case styleName
            Case "cover_dark", "section_dark", "dark"
                background = "#0B1F33"
            Case "content_tinted", "tinted", "recommendation"
                background = "#EAF0F4"
            Case Else
                background = "#F6F8FA"
                styleName = "content_light"
        End Select

        SetSlideBackgroundColor(sp, background)
        NormalizeConsultingPlaceholderGeometry(presPart, sp, styleName)

        Dim slideWidth As System.Int64 = presPart.Presentation.SlideSize.Cx.Value
        Dim slideHeight As System.Int64 = presPart.Presentation.SlideSize.Cy.Value
        Dim accent As System.String = "#1F6F8B"
        Dim accentWarm As System.String = "#D97745"

        If styleName = "cover_dark" OrElse styleName = "section_dark" Then
            AddPrimitiveShape(
                sp,
                New EmuRect With {
                    .X = CLng(slideWidth * 0.065R),
                    .Y = CLng(slideHeight * 0.255R),
                    .Width = CLng(slideWidth * 0.006R),
                    .Height = CLng(slideHeight * 0.31R)
                },
                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle,
                accent)
            AddPrimitiveShape(
                sp,
                New EmuRect With {
                    .X = CLng(slideWidth * 0.84R),
                    .Y = 0L,
                    .Width = CLng(slideWidth * 0.11R),
                    .Height = CLng(slideHeight * 0.012R)
                },
                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle,
                accentWarm)
        Else
            AddPrimitiveShape(
                sp,
                New EmuRect With {
                    .X = 0L,
                    .Y = 0L,
                    .Width = slideWidth,
                    .Height = CLng(slideHeight * 0.009R)
                },
                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle,
                accent)
        End If
    End Sub

    Private Sub SetSlideBackgroundColor(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal color As System.String)

        If sp?.Slide?.CommonSlideData Is Nothing Then Return
        Dim commonData As DocumentFormat.OpenXml.Presentation.CommonSlideData = sp.Slide.CommonSlideData
        Dim existing As DocumentFormat.OpenXml.Presentation.Background = commonData.GetFirstChild(Of DocumentFormat.OpenXml.Presentation.Background)()
        If existing IsNot Nothing Then existing.Remove()

        Dim background As New DocumentFormat.OpenXml.Presentation.Background(
            New DocumentFormat.OpenXml.Presentation.BackgroundProperties(
                New DocumentFormat.OpenXml.Drawing.SolidFill(
                    New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {
                        .Val = NormalizeHexColor(color, "#F6F8FA").TrimStart("#"c)
                    })))

        ' p:bg must precede p:spTree in p:cSld.
        commonData.PrependChild(background)
    End Sub

    Private Sub NormalizeConsultingPlaceholderGeometry(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal slideStyle As System.String)

        If presPart?.Presentation?.SlideSize Is Nothing OrElse sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return
        Dim slideWidth As System.Int64 = presPart.Presentation.SlideSize.Cx.Value
        Dim slideHeight As System.Int64 = presPart.Presentation.SlideSize.Cy.Value
        Dim styleName As System.String = If(slideStyle, System.String.Empty).Trim().ToLowerInvariant()

        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = FindFirstShapeByRole(sp, "title")
        Dim subtitleShape As DocumentFormat.OpenXml.Presentation.Shape = FindFirstShapeByRole(sp, "subtitle")
        Dim bodyShapes As New System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape)()

        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            Dim role As System.String = ResolvePlaceholderRoleForJson(ph, If(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty))
            If System.String.Equals(role, "body", System.StringComparison.OrdinalIgnoreCase) Then bodyShapes.Add(shape)
        Next

        If styleName = "cover_dark" OrElse styleName = "section_dark" Then
            If titleShape IsNot Nothing Then SetShapeRect(titleShape, New EmuRect With {
                .X = CLng(slideWidth * 0.09R), .Y = CLng(slideHeight * 0.27R),
                .Width = CLng(slideWidth * 0.78R), .Height = CLng(slideHeight * 0.22R)})
            If subtitleShape IsNot Nothing Then SetShapeRect(subtitleShape, New EmuRect With {
                .X = CLng(slideWidth * 0.09R), .Y = CLng(slideHeight * 0.52R),
                .Width = CLng(slideWidth * 0.68R), .Height = CLng(slideHeight * 0.14R)})
        Else
            If titleShape IsNot Nothing Then SetShapeRect(titleShape, New EmuRect With {
                .X = CLng(slideWidth * 0.062R), .Y = CLng(slideHeight * 0.047R),
                .Width = CLng(slideWidth * 0.876R), .Height = CLng(slideHeight * 0.145R)})

            If bodyShapes.Count = 1 Then
                SetShapeRect(bodyShapes(0), New EmuRect With {
                    .X = CLng(slideWidth * 0.062R), .Y = CLng(slideHeight * 0.235R),
                    .Width = CLng(slideWidth * 0.876R), .Height = CLng(slideHeight * 0.675R)})
            ElseIf bodyShapes.Count = 2 Then
                SetShapeRect(bodyShapes(0), New EmuRect With {
                    .X = CLng(slideWidth * 0.062R), .Y = CLng(slideHeight * 0.235R),
                    .Width = CLng(slideWidth * 0.414R), .Height = CLng(slideHeight * 0.675R)})
                SetShapeRect(bodyShapes(1), New EmuRect With {
                    .X = CLng(slideWidth * 0.524R), .Y = CLng(slideHeight * 0.235R),
                    .Width = CLng(slideWidth * 0.414R), .Height = CLng(slideHeight * 0.675R)})
            End If
        End If
    End Sub

    Private Sub SetShapeRect(
        ByVal shape As DocumentFormat.OpenXml.Presentation.Shape,
        ByVal rect As EmuRect)

        If shape Is Nothing Then Return
        If shape.ShapeProperties Is Nothing Then shape.ShapeProperties = New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        shape.ShapeProperties.Transform2D = RectTransform(rect)
    End Sub

    Private Sub ApplyConsultingPlaceholderTypography(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal slideStyle As System.String)

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return
        Dim styleName As System.String = If(slideStyle, System.String.Empty).Trim().ToLowerInvariant()
        Dim isDark As System.Boolean = styleName = "cover_dark" OrElse styleName = "section_dark" OrElse styleName = "dark"

        Dim profile As New DesignProfileJson With {
            .Palette = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        }
        FillThemeProfile(sp.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme, profile)
        Dim headingFont As System.String = If(System.String.IsNullOrWhiteSpace(profile.HeadingFont), "Aptos Display", profile.HeadingFont)
        Dim bodyFont As System.String = If(System.String.IsNullOrWhiteSpace(profile.BodyFont), "Aptos", profile.BodyFont)
        Dim titleColor As System.String = If(isDark, "#F7FAFC", "#142536")
        Dim bodyColor As System.String = If(isDark, "#DCE5EC", "#334A5E")

        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing OrElse shape.TextBody Is Nothing Then Continue For
            Dim role As System.String = ResolvePlaceholderRoleForJson(ph, If(shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty))

            Select Case role.ToLowerInvariant()
                Case "title"
                    ApplyExplicitRunStyleToShape(shape, headingFont, If(isDark, 30.0R, 24.0R), titleColor, True)
                Case "subtitle"
                    ApplyExplicitRunStyleToShape(shape, bodyFont, 15.5R, bodyColor, False)
                Case "body"
                    ApplyExplicitRunStyleToShape(shape, bodyFont, 15.0R, bodyColor, Nothing)
            End Select
        Next
    End Sub

    Private Sub ApplyExplicitRunStyleToShape(
        ByVal shape As DocumentFormat.OpenXml.Presentation.Shape,
        ByVal fontName As System.String,
        ByVal fontSize As System.Double,
        ByVal color As System.String,
        ByVal bold As System.Nullable(Of System.Boolean))

        If shape?.TextBody Is Nothing Then Return
        For Each run As DocumentFormat.OpenXml.Drawing.Run In shape.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Run)()
            If run.RunProperties Is Nothing Then run.RunProperties = New DocumentFormat.OpenXml.Drawing.RunProperties()
            run.RunProperties.FontSize = CInt(fontSize * 100.0R)
            If bold.HasValue Then run.RunProperties.Bold = bold.Value
            SetRunSolidColorOrdered(run.RunProperties, color)
            SetRunLatinFontOrdered(run.RunProperties, fontName)
        Next
    End Sub

    Private Function InsertAtAnchor(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal mode As System.String,
        ByVal anchorSlideId As UInteger,
        ByVal newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart) As UInteger

        Dim slideList As DocumentFormat.OpenXml.Presentation.SlideIdList = presPart.Presentation.SlideIdList
        Dim relId As System.String = presPart.GetIdOfPart(newSlidePart)
        Dim existing As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.SlideId) = slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)().ToList()
        Dim newId As UInteger = If(existing.Count > 0, existing.Max(Function(s) s.Id.Value) + 1UI, 256UI)
        Dim newSlide As New DocumentFormat.OpenXml.Presentation.SlideId() With {.Id = newId, .RelationshipId = relId}

        Dim normalizedMode As System.String = If(mode, "at_end").Trim().ToLowerInvariant()
        If normalizedMode = "at_end" OrElse anchorSlideId = 0UI Then
            slideList.Append(newSlide)
            Return newId
        End If

        Dim anchor As DocumentFormat.OpenXml.Presentation.SlideId = existing.FirstOrDefault(Function(s) s.Id IsNot Nothing AndAlso s.Id.Value = anchorSlideId)
        If anchor Is Nothing Then
            slideList.Append(newSlide)
        ElseIf normalizedMode = "before" Then
            anchor.InsertBeforeSelf(newSlide)
        Else
            anchor.InsertAfterSelf(newSlide)
        End If
        Return newId
    End Function


    ''' <summary>
    ''' Validates a PPTX file for OpenXML compliance errors.
    ''' </summary>
    ''' <param name="path">Full path to the PowerPoint file to validate.</param>
    ''' <returns>Error description string if validation fails; empty string if no errors found.</returns>
    ''' <remarks>Returns only the first error found for debugging purposes.</remarks>
    Function ValidatePptx(path As String) As String

        Dim ErrorString As String = ""

        Using doc As PresentationDocument = PresentationDocument.Open(path, False)

            Dim validator As New OpenXmlValidator()
            Dim errors = validator.Validate(doc)

            If Not errors.Any() Then
                Debug.WriteLine("No formal OpenXML-Errors found.")
                Return ""
            End If

            For Each err As ValidationErrorInfo In errors
                Debug.WriteLine("----------")
                Debug.WriteLine($"Part : {err.Part.Uri}")
                Debug.WriteLine($"XPath: {err.Path.XPath}")
                Debug.WriteLine($"Info : {err.Description}")
                ErrorString = $"Part: {err.Part.Uri}; XPath: {err.Path.XPath}; Info: {err.Description}"
                ' Stop after first error – sufficient for debugging
                Exit For
            Next

        End Using

        Return ErrorString

    End Function


    ' Overload: Takes PresentationPart instead of path

    ''' <summary>
    ''' Builds a deck index from a presentation part for slide lookups.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>DeckIndex with slide key and ID mappings.</returns>    
    Public Function BuildDeckIndex(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DeckIndex
        Dim idx As New DeckIndex With {
        .SlideKeyById = New Dictionary(Of String, UInteger)(),
        .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
    }
        Dim i As Integer = 0
        For Each sid In presPart.Presentation.SlideIdList.Elements(
                            Of DocumentFormat.OpenXml.Presentation.SlideId)()
            idx.IndexBySlideId(sid.Id.Value) = i
            Dim sp = CType(presPart.GetPartById(sid.RelationshipId),
                       DocumentFormat.OpenXml.Packaging.SlidePart)
            Dim key = GetSlideKey(sp, sid.Id.Value)
            idx.SlideKeyById(key) = sid.Id.Value
            i += 1
        Next
        Return idx
    End Function

    ''' <summary>
    ''' Builds a deck index from a PPTX file path.
    ''' </summary>
    ''' <param name="pptxPath">Full path to the PowerPoint file.</param>
    ''' <returns>DeckIndex with slide key and ID mappings.</returns>
    Public Function BuildDeckIndex(pptxPath As String) As DeckIndex
        Using presDoc As DocumentFormat.OpenXml.Packaging.PresentationDocument =
              DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(pptxPath, False)
            Dim presPart = presDoc.PresentationPart
            Dim idx As New DeckIndex With {
              .SlideKeyById = New Dictionary(Of String, UInteger)(),
              .IndexBySlideId = New Dictionary(Of UInteger, Integer)()
            }
            Dim i As Integer = 0
            For Each sid As DocumentFormat.OpenXml.Presentation.SlideId _
                In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
                idx.IndexBySlideId(sid.Id.Value) = i
                Dim sp = CType(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
                Dim key = GetSlideKey(sp, sid.Id.Value)
                idx.SlideKeyById(key) = sid.Id.Value
                i += 1
            Next
            Return idx
        End Using
    End Function


    ''' <summary>
    ''' Clones a template slide from the specified layout.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="layoutRelId">Layout relationship ID, URI, or name.</param>
    ''' <returns>A new slide part cloned from the layout.</returns>
    ''' <summary>
    ''' Creates a new slide that inherits the selected layout correctly.
    ''' Only placeholder instances are copied to the slide; decorative layout artwork remains on the layout
    ''' and is therefore inherited once, rather than being duplicated into the slide XML.
    ''' </summary>
    Private Function CloneTemplateSlide(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart
    ) As DocumentFormat.OpenXml.Packaging.SlidePart

        If presPart Is Nothing Then Throw New System.Exception("PresentationPart is missing.")
        If targetLayout Is Nothing Then Throw New System.Exception("No valid slide layout was resolved.")

        Dim newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart =
            presPart.AddNewPart(Of DocumentFormat.OpenXml.Packaging.SlidePart)()

        Dim sourceTree As DocumentFormat.OpenXml.Presentation.ShapeTree = targetLayout.SlideLayout?.CommonSlideData?.ShapeTree
        Dim targetTree As New DocumentFormat.OpenXml.Presentation.ShapeTree(
            New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1UI, .Name = System.String.Empty},
                New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
            New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())

        If sourceTree IsNot Nothing Then
            For Each child As DocumentFormat.OpenXml.OpenXmlElement In sourceTree.ChildElements
                If TypeOf child Is DocumentFormat.OpenXml.Presentation.Shape Then
                    Dim shp As DocumentFormat.OpenXml.Presentation.Shape = CType(child, DocumentFormat.OpenXml.Presentation.Shape)
                    Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                        shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph IsNot Nothing Then
                        Dim layoutShapeName As System.String =
                            If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
                        Dim semanticRole As System.String = ResolvePlaceholderSemanticRole(ph, layoutShapeName)

                        ' Some corporate templates implement logos and other artwork as placeholders with
                        ' relationship-backed fills (for example the VISCHER SVG logo). PowerPoint may require
                        ' the corresponding slide-level placeholder instance in order to render that artwork.
                        ' Therefore keep the placeholder and copy/rewrite its image relationship below.
                        Dim cloned As DocumentFormat.OpenXml.Presentation.Shape = CType(shp.CloneNode(True), DocumentFormat.OpenXml.Presentation.Shape)

                        ' Never purge the contents of a visual/logo placeholder. For normal editable text
                        ' placeholders, retain the formatting shell but remove the sample text.
                        If Not ShapeContainsRelationshipBackedVisual(shp) AndAlso
                           Not System.String.Equals(semanticRole, "logo", System.StringComparison.OrdinalIgnoreCase) Then
                            ClearPlaceholderTextPreserveFormatting(cloned)
                        End If

                        targetTree.Append(cloned)
                    End If
                ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.Picture Then
                    Dim pic As DocumentFormat.OpenXml.Presentation.Picture = CType(child, DocumentFormat.OpenXml.Presentation.Picture)
                    If pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape IsNot Nothing Then
                        ' Preserve the complete picture-placeholder shell. If it contains a relationship-backed
                        ' default image, CopyLayoutImagesToSlide will copy the image part and rewrite the r:id.
                        Dim clonedPic As DocumentFormat.OpenXml.Presentation.Picture = CType(pic.CloneNode(True), DocumentFormat.OpenXml.Presentation.Picture)
                        targetTree.Append(clonedPic)
                    End If
                ElseIf TypeOf child Is DocumentFormat.OpenXml.Presentation.GraphicFrame Then
                    Dim gf As DocumentFormat.OpenXml.Presentation.GraphicFrame = CType(child, DocumentFormat.OpenXml.Presentation.GraphicFrame)
                    If gf.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape IsNot Nothing Then
                        ' Preserve the placeholder shell but not layout-local chart/table relationships.
                        Dim clonedGf As DocumentFormat.OpenXml.Presentation.GraphicFrame = CType(gf.CloneNode(True), DocumentFormat.OpenXml.Presentation.GraphicFrame)
                        Dim graphicData As DocumentFormat.OpenXml.Drawing.GraphicData = clonedGf.Graphic?.GraphicData
                        If graphicData IsNot Nothing Then graphicData.RemoveAllChildren()
                        targetTree.Append(clonedGf)
                    End If
                End If
            Next
        End If

        Dim commonData As New DocumentFormat.OpenXml.Presentation.CommonSlideData(targetTree)
        Dim newSlide As New DocumentFormat.OpenXml.Presentation.Slide(commonData)

        If targetLayout.SlideLayout?.ColorMapOverride IsNot Nothing Then
            newSlide.ColorMapOverride = CType(targetLayout.SlideLayout.ColorMapOverride.CloneNode(True), DocumentFormat.OpenXml.Presentation.ColorMapOverride)
        End If

        newSlidePart.Slide = newSlide

        ' Placeholder artwork may reference images owned by the layout. Copy those image parts to the
        ' slide and rewrite every relationship-backed reference (including SVG extension elements) before
        ' attaching the layout. This preserves branded placeholder artwork without broken rIds.
        CopyLayoutImagesToSlide(targetLayout, newSlidePart)

        newSlidePart.AddPart(targetLayout)
        newSlidePart.Slide.Save()
        Return newSlidePart
    End Function

    ''' <summary>
    ''' Legacy-compatible overload. It now resolves a layout and then uses the correct inheritance-based cloning path.
    ''' </summary>
    Private Function CloneTemplateSlide(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal layoutSelector As System.String
    ) As DocumentFormat.OpenXml.Packaging.SlidePart

        Dim targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart = ResolveLayout(presPart, layoutSelector)
        If targetLayout Is Nothing Then
            targetLayout = PickDefaultLayout(presPart)
        End If
        Return CloneTemplateSlide(presPart, targetLayout)
    End Function

    Private Sub ClearPlaceholderTextPreserveFormatting(ByVal shp As DocumentFormat.OpenXml.Presentation.Shape)
        If shp Is Nothing Then Return
        If shp.TextBody Is Nothing Then
            shp.TextBody = New DocumentFormat.OpenXml.Presentation.TextBody(
                New DocumentFormat.OpenXml.Drawing.BodyProperties(),
                New DocumentFormat.OpenXml.Drawing.ListStyle(),
                New DocumentFormat.OpenXml.Drawing.Paragraph(New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()))
            Return
        End If

        Dim oldTb As DocumentFormat.OpenXml.Presentation.TextBody = shp.TextBody
        Dim newTb As New DocumentFormat.OpenXml.Presentation.TextBody()
        If oldTb.BodyProperties IsNot Nothing Then
            newTb.Append(CType(oldTb.BodyProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.BodyProperties))
        Else
            newTb.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        End If
        If oldTb.ListStyle IsNot Nothing Then
            newTb.Append(CType(oldTb.ListStyle.CloneNode(True), DocumentFormat.OpenXml.Drawing.ListStyle))
        Else
            newTb.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())
        End If

        Dim templateParagraph As DocumentFormat.OpenXml.Drawing.Paragraph = oldTb.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().FirstOrDefault()
        Dim emptyParagraph As New DocumentFormat.OpenXml.Drawing.Paragraph()
        If templateParagraph?.ParagraphProperties IsNot Nothing Then
            emptyParagraph.Append(CType(templateParagraph.ParagraphProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.ParagraphProperties))
        End If
        Dim templateRun As DocumentFormat.OpenXml.Drawing.Run = templateParagraph?.Elements(Of DocumentFormat.OpenXml.Drawing.Run)().FirstOrDefault()
        If templateRun?.RunProperties IsNot Nothing Then
            emptyParagraph.Append(New DocumentFormat.OpenXml.Drawing.Run(
                CType(templateRun.RunProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.RunProperties),
                New DocumentFormat.OpenXml.Drawing.Text(System.String.Empty)))
        End If
        Dim templateEndProperties As DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties =
            If(templateParagraph IsNot Nothing,
               templateParagraph.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties)(),
               Nothing)
        If templateEndProperties IsNot Nothing Then
            emptyParagraph.Append(CType(templateEndProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties))
        Else
            emptyParagraph.Append(New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties())
        End If
        newTb.Append(emptyParagraph)
        shp.TextBody = newTb
    End Sub

    Private Function ResolveLayoutFromAction(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal action As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        Dim relId As System.String = System.String.Empty
        Dim uri As System.String = System.String.Empty
        Dim name As System.String = System.String.Empty
        Dim master As System.String = System.String.Empty
        Dim tmp As System.Text.Json.JsonElement

        If action.TryGetProperty("layoutRelId", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then relId = tmp.GetString()
        If action.TryGetProperty("layoutId", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then uri = tmp.GetString()
        If action.TryGetProperty("layoutName", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then name = tmp.GetString()

        If action.TryGetProperty("layoutKey", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Object Then
            Dim k As System.Text.Json.JsonElement
            If tmp.TryGetProperty("relId", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then relId = k.GetString()
            If tmp.TryGetProperty("layoutRelId", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then relId = k.GetString()

            ' Canonical key is "uri". Accept "layoutId" as a compatibility alias because
            ' LLMs naturally mirror the metadata property name. Ignoring it used to make
            ' every action silently fall back to PickDefaultLayout().
            If tmp.TryGetProperty("uri", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then uri = k.GetString()
            If tmp.TryGetProperty("layoutId", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then uri = k.GetString()

            If tmp.TryGetProperty("name", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then name = k.GetString()
            If tmp.TryGetProperty("layoutName", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then name = k.GetString()
            If tmp.TryGetProperty("master", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then master = k.GetString()
            If tmp.TryGetProperty("masterId", k) AndAlso k.ValueKind = System.Text.Json.JsonValueKind.String Then master = k.GetString()
        End If

        Dim matches As New System.Collections.Generic.List(Of DocumentFormat.OpenXml.Packaging.SlideLayoutPart)()
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            Dim smName As System.String = GetMasterName(sm)
            Dim smId As System.String = If(sm.Uri IsNot Nothing, sm.Uri.ToString(), System.String.Empty)
            If Not System.String.IsNullOrWhiteSpace(master) AndAlso
               Not System.String.Equals(master, smName, System.StringComparison.OrdinalIgnoreCase) AndAlso
               Not System.String.Equals(master, smId, System.StringComparison.OrdinalIgnoreCase) Then Continue For

            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim rid As System.String = System.String.Empty
                Try
                    rid = sm.GetIdOfPart(lp)
                Catch ex As System.Exception
                    rid = System.String.Empty
                End Try

                If Not System.String.IsNullOrWhiteSpace(uri) AndAlso lp.Uri IsNot Nothing AndAlso
                   System.String.Equals(lp.Uri.ToString(), uri, System.StringComparison.OrdinalIgnoreCase) Then Return lp

                If Not System.String.IsNullOrWhiteSpace(name) AndAlso
                   System.String.Equals(GetLayoutName(lp), name, System.StringComparison.OrdinalIgnoreCase) Then matches.Add(lp)

                If Not System.String.IsNullOrWhiteSpace(relId) AndAlso
                   System.String.Equals(rid, relId, System.StringComparison.OrdinalIgnoreCase) Then matches.Add(lp)
            Next
        Next

        matches = matches.Distinct().ToList()
        If matches.Count = 1 Then Return matches(0)
        If matches.Count > 1 Then
            Throw New System.Exception("The requested layout selector is ambiguous across slide masters. Include layoutKey.uri or layoutKey.masterId.")
        End If

        Dim hadSelector As System.Boolean = Not System.String.IsNullOrWhiteSpace(relId) OrElse Not System.String.IsNullOrWhiteSpace(uri) OrElse Not System.String.IsNullOrWhiteSpace(name)
        If hadSelector Then
            Throw New System.Exception("The requested layout does not exist in the presentation metadata.")
        End If

        Return PickDefaultLayout(presPart)
    End Function

    Private Function CloneExistingSlideAsTemplate(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal source As DocumentFormat.OpenXml.Packaging.SlidePart) As DocumentFormat.OpenXml.Packaging.SlidePart

        If presPart Is Nothing OrElse source Is Nothing Then Throw New System.Exception("The source slide for cloning could not be resolved.")
        Dim target As DocumentFormat.OpenXml.Packaging.SlidePart = presPart.AddNewPart(Of DocumentFormat.OpenXml.Packaging.SlidePart)()
        target.Slide = CType(source.Slide.CloneNode(True), DocumentFormat.OpenXml.Presentation.Slide)

        For Each pair In source.Parts
            If TypeOf pair.OpenXmlPart Is DocumentFormat.OpenXml.Packaging.NotesSlidePart Then Continue For
            Try
                target.AddPart(pair.OpenXmlPart, pair.RelationshipId)
            Catch ex As System.Exception
                Try
                    target.AddPart(pair.OpenXmlPart)
                Catch innerEx As System.Exception
                    System.Diagnostics.Debug.WriteLine("Could not copy slide relationship " & pair.RelationshipId & ": " & innerEx.Message)
                End Try
            End Try
        Next

        For Each externalRel In source.ExternalRelationships
            Try
                target.AddExternalRelationship(externalRel.RelationshipType, externalRel.Uri, externalRel.Id)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Could not copy external slide relationship: " & ex.Message)
            End Try
        Next

        target.Slide.Save()
        Return target
    End Function

    Private Function FindRuntimeSampleSlideForLayout(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal targetLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
        ByVal runtimeSampleKeys As System.Collections.Generic.HashSet(Of System.String)) As DocumentFormat.OpenXml.Packaging.SlidePart

        If presPart Is Nothing OrElse targetLayout Is Nothing OrElse runtimeSampleKeys Is Nothing OrElse runtimeSampleKeys.Count = 0 Then Return Nothing
        Dim index As DeckIndex = BuildDeckIndex(presPart)

        For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In
            presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()

            Dim sp As DocumentFormat.OpenXml.Packaging.SlidePart = TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
            If sp Is Nothing OrElse sp.SlideLayoutPart Is Nothing Then Continue For
            If sp.SlideLayoutPart.Uri Is Nothing OrElse targetLayout.Uri Is Nothing OrElse
               Not System.String.Equals(sp.SlideLayoutPart.Uri.ToString(), targetLayout.Uri.ToString(), System.StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim key As System.String = GetSlideKey(sp, sid.Id.Value)
            If runtimeSampleKeys.Contains(key) Then Return sp
        Next
        Return Nothing
    End Function

    Private Sub ClearAllTemplateCloneEditableText(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart)

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return
        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In
            sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing OrElse shape.TextBody Is Nothing Then Continue For
            Dim semanticRole As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shape)
            If semanticRole = "footer" OrElse
               semanticRole = "slide_number" OrElse
               semanticRole = "logo" OrElse
               semanticRole = "hero_image" OrElse
               ShapeContainsRelationshipBackedVisual(shape) Then
                Continue For
            End If
            ClearPlaceholderTextPreserveFormatting(shape)
        Next
        sp.Slide.Save()
    End Sub

    Private Function ShouldPreferSampleSlideClone(
        ByVal layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart) As System.Boolean

        If layoutPart Is Nothing Then Return False
        Dim role As System.String = ClassifyLayoutSemanticRole(layoutPart)
        Return System.String.Equals(role, "cover", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(role, "closing", System.StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function FindSlidePartByKey(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal slideKey As System.String) As DocumentFormat.OpenXml.Packaging.SlidePart

        If presPart Is Nothing OrElse System.String.IsNullOrWhiteSpace(slideKey) Then Return Nothing
        Dim index As DeckIndex = BuildDeckIndex(presPart)
        If Not index.SlideKeyById.ContainsKey(slideKey) Then Return Nothing
        Dim slideId As UInteger = index.SlideKeyById(slideKey)
        For Each sid As DocumentFormat.OpenXml.Presentation.SlideId In presPart.Presentation.SlideIdList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
            If sid.Id IsNot Nothing AndAlso sid.Id.Value = slideId Then
                Return TryCast(presPart.GetPartById(sid.RelationshipId), DocumentFormat.OpenXml.Packaging.SlidePart)
            End If
        Next
        Return Nothing
    End Function

    Private Sub ClearShapeTextByIds(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal action As System.Text.Json.JsonElement)
        If sp Is Nothing Then Return

        Dim clearAll As System.Boolean = False
        Dim clearAllEl As System.Text.Json.JsonElement
        If action.TryGetProperty("clearAllEditableText", clearAllEl) AndAlso
           (clearAllEl.ValueKind = System.Text.Json.JsonValueKind.True OrElse clearAllEl.ValueKind = System.Text.Json.JsonValueKind.False) Then
            clearAll = clearAllEl.GetBoolean()
        End If

        Dim ids As New System.Collections.Generic.HashSet(Of UInteger)()
        Dim idsEl As System.Text.Json.JsonElement
        If action.TryGetProperty("clearShapeIds", idsEl) AndAlso idsEl.ValueKind = System.Text.Json.JsonValueKind.Array Then
            For Each idEl As System.Text.Json.JsonElement In idsEl.EnumerateArray()
                If idEl.ValueKind = System.Text.Json.JsonValueKind.Number Then
                    Try
                        ids.Add(CUInt(idEl.GetInt64()))
                    Catch ex As System.Exception
                    End Try
                End If
            Next
        End If

        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim shapeId As DocumentFormat.OpenXml.UInt32Value = shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id
            If shapeId Is Nothing Then Continue For

            Dim shouldClear As System.Boolean = ids.Contains(shapeId.Value)
            If clearAll AndAlso shp.TextBody IsNot Nothing Then
                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                Dim isMetadata As System.Boolean = ph?.Type IsNot Nothing AndAlso
                    (ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer OrElse
                     ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime OrElse
                     ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber OrElse
                     ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Header)
                If Not isMetadata Then shouldClear = True
            End If

            If shouldClear Then ClearPlaceholderTextPreserveFormatting(shp)
        Next
    End Sub


    ''' <summary>
    ''' Resolves a layout by relationship ID, URI, or name.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="requested">The layout identifier to resolve.</param>
    ''' <returns>The matching SlideLayoutPart or Nothing if not found.</returns>
    Private Function ResolveLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    requested As System.String
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        If System.String.IsNullOrWhiteSpace(requested) Then Return Nothing

        Dim req = requested.Trim()

        ' 1) Try as exact relId
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim rid As System.String = System.String.Empty
                Try : rid = sm.GetIdOfPart(lp) : Catch : End Try
                If Not System.String.IsNullOrEmpty(rid) AndAlso
               System.String.Equals(rid, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        ' 2) Try by URI string
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim u = lp.Uri?.ToString()
                If Not System.String.IsNullOrEmpty(u) AndAlso
               System.String.Equals(u, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        ' 3) Try by human-readable layout name (e.g., "Title Slide" / "Titel")
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim name As System.String = GetLayoutName(lp)
                If Not System.String.IsNullOrEmpty(name) AndAlso
               System.String.Equals(name, req, System.StringComparison.OrdinalIgnoreCase) Then
                    Return lp
                End If
            Next
        Next

        Return Nothing
    End Function

    ''' <summary>
    ''' Picks a cover-like layout (Title + Subtitle, preferably without Body).
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>A suitable SlideLayoutPart or Nothing.</returns>
    Private Function PickCoverLikeLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim li = AnalyzeLayoutPlaceholders(lp)
                ' Typical title slide: Title + Subtitle; often NO Body placeholder
                If (li.HasTitle OrElse li.HasCenteredTitle) AndAlso li.HasSubTitle AndAlso Not li.HasBody Then
                    Return lp
                End If
            Next
        Next

        ' next best: Title + Subtitle, even if Body exists
        For Each sm As DocumentFormat.OpenXml.Packaging.SlideMasterPart In presPart.SlideMasterParts
            For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In sm.SlideLayoutParts
                Dim li = AnalyzeLayoutPlaceholders(lp)
                If (li.HasTitle OrElse li.HasCenteredTitle) AndAlso li.HasSubTitle Then
                    Return lp
                End If
            Next
        Next

        Return Nothing
    End Function


    ''' <summary>
    ''' Picks a default layout with Title and Body placeholders.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <returns>A suitable SlideLayoutPart.</returns>
    ''' <exception cref="System.Exception">Thrown when no layout is available.</exception>
    Private Function PickDefaultLayout(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart
) As DocumentFormat.OpenXml.Packaging.SlideLayoutPart

        Dim firstMaster As DocumentFormat.OpenXml.Packaging.SlideMasterPart =
        presPart.SlideMasterParts.FirstOrDefault()
        If firstMaster Is Nothing Then
            Throw New System.Exception("No SlideMasterPart found in the presentation.")
        End If

        ' Prefer a layout that has both Title and Body placeholders
        For Each lp As DocumentFormat.OpenXml.Packaging.SlideLayoutPart In firstMaster.SlideLayoutParts
            Dim hasTitle As Boolean = False
            Dim hasBody As Boolean = False

            If lp.SlideLayout IsNot Nothing AndAlso
           lp.SlideLayout.CommonSlideData IsNot Nothing AndAlso
           lp.SlideLayout.CommonSlideData.ShapeTree IsNot Nothing Then

                Dim shapes =
                lp.SlideLayout.CommonSlideData.ShapeTree.
                    Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

                For Each sh As DocumentFormat.OpenXml.Presentation.Shape In shapes
                    Dim nv As DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties =
                    sh.NonVisualShapeProperties
                    If nv Is Nothing Then Continue For

                    Dim app As DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties =
                    nv.ApplicationNonVisualDrawingProperties
                    If app Is Nothing Then Continue For

                    Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                    app.PlaceholderShape
                    If ph Is Nothing OrElse ph.Type Is Nothing Then Continue For

                    Dim t As DocumentFormat.OpenXml.Presentation.PlaceholderValues = ph.Type.Value
                    If t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse
                   t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle Then
                        hasTitle = True
                    ElseIf t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then
                        hasBody = True
                    End If

                    If hasTitle AndAlso hasBody Then
                        Return lp
                    End If
                Next
            End If
        Next

        ' Fallback: first available layout
        Dim anyLayout As DocumentFormat.OpenXml.Packaging.SlideLayoutPart =
        firstMaster.SlideLayoutParts.FirstOrDefault()
        If anyLayout Is Nothing Then
            Throw New System.Exception("No SlideLayoutPart available to create a new slide.")
        End If
        Return anyLayout
    End Function



    ''' <summary>
    ''' Copies image parts from a layout to a slide and rewrites embed IDs.
    ''' </summary>
    ''' <param name="layoutPart">The source layout part.</param>
    ''' <param name="slidePart">The target slide part.</param>
    Private Sub CopyLayoutImagesToSlide(
        ByVal layoutPart As DocumentFormat.OpenXml.Packaging.SlideLayoutPart,
        ByVal slidePart As DocumentFormat.OpenXml.Packaging.SlidePart)

        If layoutPart Is Nothing OrElse slidePart Is Nothing OrElse slidePart.Slide Is Nothing Then Return

        ' old layout relationship id -> new slide relationship id
        Dim idMap As New System.Collections.Generic.Dictionary(Of System.String, System.String)(
            System.StringComparer.OrdinalIgnoreCase)

        ' Clone image parts used by relationship-backed placeholders. Copying all layout ImageParts is
        ' intentional: it keeps the operation simple and guarantees that SVG/PNG/JPEG placeholder artwork
        ' can be resolved even when the relationship occurs inside an extension element.
        For Each img As DocumentFormat.OpenXml.Packaging.ImagePart In layoutPart.ImageParts
            Try
                Dim oldId As System.String = layoutPart.GetIdOfPart(img)
                If System.String.IsNullOrWhiteSpace(oldId) Then Continue For

                Dim newImg As DocumentFormat.OpenXml.Packaging.ImagePart = slidePart.AddImagePart(img.ContentType)
                Using src As System.IO.Stream = img.GetStream(System.IO.FileMode.Open, System.IO.FileAccess.Read),
                      dst As System.IO.Stream = newImg.GetStream(System.IO.FileMode.Create, System.IO.FileAccess.Write)
                    src.CopyTo(dst)
                End Using

                idMap(oldId) = slidePart.GetIdOfPart(newImg)
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("Could not copy layout image to generated slide: " & ex.Message)
            End Try
        Next

        If idMap.Count = 0 Then Return

        ' Do not limit rewriting to a:blip/@r:embed. Modern Office SVG artwork is commonly stored as
        ' <asvg:svgBlip r:embed="..."> inside an extension list. Walk the complete cloned slide tree and
        ' rewrite every existing relationship-valued attribute whose value appears in our image map.
        Dim elements As New System.Collections.Generic.List(Of DocumentFormat.OpenXml.OpenXmlElement)()
        elements.Add(slidePart.Slide)
        elements.AddRange(slidePart.Slide.Descendants())

        For Each element As DocumentFormat.OpenXml.OpenXmlElement In elements
            Dim attributes As System.Collections.Generic.List(Of DocumentFormat.OpenXml.OpenXmlAttribute) =
                element.GetAttributes().ToList()

            For Each attribute As DocumentFormat.OpenXml.OpenXmlAttribute In attributes
                If System.String.IsNullOrWhiteSpace(attribute.Value) OrElse Not idMap.ContainsKey(attribute.Value) Then
                    Continue For
                End If

                Dim isRelationshipAttribute As System.Boolean =
                    System.String.Equals(
                        attribute.NamespaceUri,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
                        System.StringComparison.OrdinalIgnoreCase) OrElse
                    System.String.Equals(attribute.Prefix, "r", System.StringComparison.OrdinalIgnoreCase)

                If Not isRelationshipAttribute Then Continue For

                Try
                    element.SetAttribute(
                        New DocumentFormat.OpenXml.OpenXmlAttribute(
                            attribute.Prefix,
                            attribute.LocalName,
                            attribute.NamespaceUri,
                            idMap(attribute.Value)))
                Catch ex As System.Exception
                    System.Diagnostics.Debug.WriteLine(
                        "Could not rewrite layout image relationship '" & attribute.Value & "': " & ex.Message)
                End Try
            Next
        Next
    End Sub


    ''' <summary>
    ''' Clears sample text from title and body placeholders in a cloned slide.
    ''' </summary>
    ''' <param name="sld">The slide to purge.</param>
    Private Sub PurgeLayoutSampleText(sld As DocumentFormat.OpenXml.Presentation.Slide)

        ' only Title / CenteredTitle / Body placeholders get wiped
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape _
        In sld.CommonSlideData.ShapeTree.
               Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                 ApplicationNonVisualDrawingProperties?.
                 PlaceholderShape
            If ph Is Nothing Then Continue For

            Dim t As DocumentFormat.OpenXml.Presentation.PlaceholderValues? = Nothing
            If ph.Type IsNot Nothing Then t = ph.Type.Value

            If t Is Nothing _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body _
           OrElse t = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object Then

                ' wipe existing content
                shp.TextBody?.Remove()

                ' insert minimal, valid skeleton
                shp.Append(New DocumentFormat.OpenXml.Presentation.TextBody(
                    New DocumentFormat.OpenXml.Drawing.BodyProperties(),
                    New DocumentFormat.OpenXml.Drawing.ListStyle(),
                    New DocumentFormat.OpenXml.Drawing.Paragraph(
                        New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties())))
            End If
        Next
    End Sub


    ''' <summary>
    ''' Ensures the presentation has a SlideIdList element.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    Private Sub EnsureSlideIdList(presPart As DocumentFormat.OpenXml.Packaging.PresentationPart)
        Dim pres = presPart.Presentation
        If pres.SlideIdList IsNot Nothing Then Exit Sub

        Dim sldIdList As New DocumentFormat.OpenXml.Presentation.SlideIdList()

        ' Find the correct insertion index:
        ' Order (simplified): SlideMasterIdList?, NotesMasterIdList?, HandoutMasterIdList?, SlideIdList?, SlideSize?, NotesSize? ...
        Dim children = pres.ChildElements
        Dim insertIndex As Integer = children.Count ' default to end, then adjust

        ' Prefer to insert BEFORE SlideSize or NotesSize if present
        For i As Integer = 0 To children.Count - 1
            If TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.SlideSize _
        OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.NotesSize Then
                insertIndex = i
                Exit For
            End If
        Next

        ' If we didn't find sizes, place right after SlideMasterIdList / NotesMasterIdList / HandoutMasterIdList if any
        If insertIndex = children.Count Then
            Dim afterIndex As Integer = -1
            For i As Integer = 0 To children.Count - 1
                If TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.SlideMasterIdList _
            OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.NotesMasterIdList _
            OrElse TypeOf children(i) Is DocumentFormat.OpenXml.Presentation.HandoutMasterIdList Then
                    afterIndex = i
                End If
            Next
            insertIndex = If(afterIndex >= 0, afterIndex + 1, 0)
        End If

        pres.InsertAt(sldIdList, insertIndex)
        pres.Save()
    End Sub


    ''' <summary>
    ''' Inserts a new slide after the specified anchor slide or at the end.
    ''' </summary>
    ''' <param name="presPart">The presentation part.</param>
    ''' <param name="anchorSlideId">The slide ID to insert after; 0 to append at end.</param>
    ''' <param name="newSlidePart">The new slide part to insert.</param>
    ''' <returns>The new slide ID assigned.</returns>
    Private Function InsertAfter(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    anchorSlideId As UInteger,
    newSlidePart As DocumentFormat.OpenXml.Packaging.SlidePart
) As UInteger

        Dim slideList = presPart.Presentation.SlideIdList
        Dim relId = presPart.GetIdOfPart(newSlidePart)

        ' Find existing SlideId elements
        Dim existing = slideList.Elements(Of DocumentFormat.OpenXml.Presentation.SlideId)()
        Dim newId As UInteger

        If existing.Any() Then
            newId = existing.Max(Function(s) s.Id.Value) + 1UI
        Else
            newId = 256UI   ' First slide
        End If

        Dim newSlide = New DocumentFormat.OpenXml.Presentation.SlideId() With {
      .Id = newId,
      .RelationshipId = relId
    }

        ' If anchorSlideId = 0, always append at the end
        If anchorSlideId = 0UI Then
            slideList.Append(newSlide)
        Else
            ' Otherwise insert after the specified anchor
            Dim anchor = existing.FirstOrDefault(Function(s) s.Id.Value = anchorSlideId)
            If anchor Is Nothing Then
                slideList.Append(newSlide)
            Else
                anchor.InsertAfterSelf(newSlide)
            End If
        End If

        Return newId
    End Function


    ''' <summary>
    ''' Sets the title text in a slide's title placeholder shape.
    ''' </summary>
    ''' <param name="sp">The slide part containing the title shape.</param>
    ''' <param name="text">The title text to insert.</param>
    ''' <param name="el">JSON element containing style properties.</param>
    ''' <remarks>
    ''' Search priority: 1) Explicit Title/CenteredTitle placeholder, 2) Placeholder with index=0, 
    ''' 3) Shape name containing "title" (case-insensitive).
    ''' </remarks>
    Private Sub SetTitle(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal text As System.String,
        ByVal el As System.Text.Json.JsonElement)

        Dim shapes = sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = shapes.FirstOrDefault(Function(shp)
                                                                                                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                                                                                Return ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
                                                                                                       (ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title OrElse ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle)
                                                                                            End Function)

        If titleShape Is Nothing Then
            titleShape = shapes.FirstOrDefault(Function(shp)
                                                   Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                                   Return ph IsNot Nothing AndAlso ph.Index IsNot Nothing AndAlso ph.Index.Value = 0UI
                                               End Function)
        End If
        If titleShape Is Nothing Then Return

        SetShapeSingleTextPreserveStyle(titleShape, text, el, forceNoBullet:=False)
        ApplyGeneratedPlaceholderTextFit(sp, titleShape, text)
        sp.Slide.Save()
    End Sub

    Private Sub SetBullets(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement)

        Dim target As DocumentFormat.OpenXml.Presentation.Shape = Nothing
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If IsBodyLikePlaceholder(ph) Then
                target = shp
                Exit For
            End If
        Next
        If target Is Nothing Then Return
        SetShapeBulletsPreserveStyle(target, el)
        sp.Slide.Save()
    End Sub

    Private Sub SetShapeSingleTextPreserveStyle(
        ByVal targetShape As DocumentFormat.OpenXml.Presentation.Shape,
        ByVal text As System.String,
        ByVal el As System.Text.Json.JsonElement,
        ByVal forceNoBullet As System.Boolean)

        If targetShape Is Nothing Then Return
        Dim oldTb As DocumentFormat.OpenXml.Presentation.TextBody = targetShape.TextBody
        Dim newTb As New DocumentFormat.OpenXml.Presentation.TextBody()
        AppendPreservedTextBodyShell(newTb, oldTb)

        Dim templateP As DocumentFormat.OpenXml.Drawing.Paragraph = oldTb?.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().FirstOrDefault()
        Dim pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = Nothing
        If templateP?.ParagraphProperties IsNot Nothing Then
            pPr = CType(templateP.ParagraphProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.ParagraphProperties)
        Else
            pPr = New DocumentFormat.OpenXml.Drawing.ParagraphProperties()
        End If

        If forceNoBullet Then
            RemoveBulletDefinition(pPr)
            pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        End If
        ApplyParagraphStyleOverrides(pPr, el)

        Dim rp As DocumentFormat.OpenXml.Drawing.RunProperties = CloneTemplateRunProperties(templateP)
        ApplyRunStyleOverrides(rp, el)
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
        If pPr.HasChildren OrElse pPr.Level IsNot Nothing OrElse pPr.Alignment IsNot Nothing Then para.Append(pPr)
        para.Append(New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(If(text, System.String.Empty))))
        If templateP IsNot Nothing Then
            Dim templateEndProperties As DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties =
                templateP.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties)()
            If templateEndProperties IsNot Nothing Then
                para.Append(CType(templateEndProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties))
            End If
        End If
        newTb.Append(para)
        targetShape.TextBody = newTb
    End Sub

    Private Sub ApplyGeneratedPlaceholderTextFit(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal targetShape As DocumentFormat.OpenXml.Presentation.Shape,
        ByVal text As System.String)

        If sp Is Nothing OrElse targetShape Is Nothing OrElse targetShape.TextBody Is Nothing OrElse System.String.IsNullOrWhiteSpace(text) Then Return

        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            targetShape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        Dim semanticRole As System.String =
            If(ph IsNot Nothing, ResolveEffectivePlaceholderSemanticRole(sp, targetShape), System.String.Empty)

        Dim style As TextStyleJson = ExtractEffectiveTextStyleForShape(sp, targetShape)
        If style Is Nothing OrElse Not style.FontSize.HasValue OrElse style.FontSize.Value <= 0.0R Then Return

        ' On a cover, preserve the native subtitle hierarchy. If the generated subtitle is
        ' longer than the sample text, first use the available vertical gap below the subtitle
        ' rather than shrinking a 28 pt corporate subtitle into tiny caption text.
        If System.String.Equals(semanticRole, "subtitle", System.StringComparison.OrdinalIgnoreCase) AndAlso
           sp.SlideLayoutPart IsNot Nothing AndAlso
           System.String.Equals(ClassifyLayoutSemanticRole(sp.SlideLayoutPart), "cover", System.StringComparison.OrdinalIgnoreCase) Then
            ExpandCoverSubtitlePlaceholder(sp, targetShape)
        End If

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, targetShape)
        If xfrm Is Nothing Then Return
        Dim rect As EmuRect = TransformToRect(xfrm)

        rect.Width = CLng(rect.Width * 0.94R)
        rect.Height = CLng(rect.Height * 0.9R)

        Dim desired As System.Double = style.FontSize.Value
        Dim fitted As System.Double = FitPrimitiveFontSize(rect, text, desired)

        Dim minimum As System.Double = System.Math.Max(9.0R, desired * 0.65R)
        If System.String.Equals(semanticRole, "subtitle", System.StringComparison.OrdinalIgnoreCase) AndAlso
           sp.SlideLayoutPart IsNot Nothing AndAlso
           System.String.Equals(ClassifyLayoutSemanticRole(sp.SlideLayoutPart), "cover", System.StringComparison.OrdinalIgnoreCase) Then
            minimum = System.Math.Max(20.0R, desired * 0.72R)
        ElseIf System.String.Equals(semanticRole, "title", System.StringComparison.OrdinalIgnoreCase) Then
            minimum = System.Math.Max(18.0R, desired * 0.72R)
        End If

        fitted = System.Math.Max(fitted, minimum)
        fitted = System.Math.Min(fitted, desired)
        If fitted >= desired - 0.25R Then Return

        For Each run As DocumentFormat.OpenXml.Drawing.Run In targetShape.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Run)()
            If run.RunProperties Is Nothing Then run.RunProperties = New DocumentFormat.OpenXml.Drawing.RunProperties()
            run.RunProperties.FontSize = CInt(fitted * 100.0R)
        Next

        For Each endProperties As DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties In
            targetShape.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties)()
            endProperties.FontSize = CInt(fitted * 100.0R)
        Next
    End Sub

    Private Sub ExpandCoverSubtitlePlaceholder(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal targetShape As DocumentFormat.OpenXml.Presentation.Shape)

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing OrElse targetShape Is Nothing Then Return
        Dim effective As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, targetShape)
        If effective Is Nothing OrElse effective.Offset Is Nothing OrElse effective.Extents Is Nothing Then Return

        Dim current As EmuRect = TransformToRect(effective)
        If current.Width <= 0 OrElse current.Height <= 0 Then Return

        Dim slideHeight As Long = If(sp.SlideLayoutPart?.SlideMasterPart?.SlideMaster IsNot Nothing AndAlso
                                    sp.SlideLayoutPart.SlideMasterPart.SlideMaster.CommonSlideData IsNot Nothing,
                                    6858000L,
                                    6858000L)
        Dim nextTop As Long = slideHeight

        For Each other As DocumentFormat.OpenXml.Presentation.Shape In
            sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

            If System.Object.ReferenceEquals(other, targetShape) Then Continue For
            Dim oph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
                other.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If oph Is Nothing Then Continue For
            Dim otherName As System.String = If(other.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
            Dim role As System.String = ResolvePlaceholderSemanticRole(oph, otherName)
            If role <> "presenter" AndAlso role <> "presenter_role" AndAlso role <> "location" AndAlso role <> "date" Then Continue For

            Dim ox As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, other)
            If ox Is Nothing Then Continue For
            Dim r As EmuRect = TransformToRect(ox)
            If r.Y <= current.Y Then Continue For

            Dim overlapLeft As Long = System.Math.Max(current.X, r.X)
            Dim overlapRight As Long = System.Math.Min(current.X + current.Width, r.X + r.Width)
            If overlapRight <= overlapLeft Then Continue For
            nextTop = System.Math.Min(nextTop, r.Y)
        Next

        Dim gap As Long = CLng(slideHeight * 0.02R)
        Dim maximumHeight As Long = nextTop - current.Y - gap
        maximumHeight = System.Math.Min(maximumHeight, CLng(slideHeight * 0.16R))
        If maximumHeight <= current.Height Then Return

        Dim direct As DocumentFormat.OpenXml.Drawing.Transform2D = targetShape.ShapeProperties?.Transform2D
        If direct Is Nothing Then
            If targetShape.ShapeProperties Is Nothing Then targetShape.ShapeProperties = New DocumentFormat.OpenXml.Presentation.ShapeProperties()
            direct = CType(effective.CloneNode(True), DocumentFormat.OpenXml.Drawing.Transform2D)
            targetShape.ShapeProperties.Transform2D = direct
        End If
        If direct.Extents Is Nothing Then direct.Extents = New DocumentFormat.OpenXml.Drawing.Extents()
        direct.Extents.Cx = current.Width
        direct.Extents.Cy = maximumHeight
    End Sub

    Private Sub SetShapeBulletsPreserveStyle(
        ByVal targetShape As DocumentFormat.OpenXml.Presentation.Shape,
        ByVal el As System.Text.Json.JsonElement)

        If targetShape Is Nothing Then Return
        Dim oldTb As DocumentFormat.OpenXml.Presentation.TextBody = targetShape.TextBody
        Dim templates As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Drawing.Paragraph) =
            If(oldTb IsNot Nothing, oldTb.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().ToList(), New System.Collections.Generic.List(Of DocumentFormat.OpenXml.Drawing.Paragraph)())

        Dim newTb As New DocumentFormat.OpenXml.Presentation.TextBody()
        AppendPreservedTextBodyShell(newTb, oldTb)

        Dim bulletsEl As System.Text.Json.JsonElement
        If Not el.TryGetProperty("bullets", bulletsEl) OrElse bulletsEl.ValueKind <> System.Text.Json.JsonValueKind.Array Then
            newTb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()))
            targetShape.TextBody = newTb
            Return
        End If

        Dim added As System.Int32 = 0
        For Each b As System.Text.Json.JsonElement In bulletsEl.EnumerateArray()
            Dim text As System.String = System.String.Empty
            Dim level As System.Int32 = 0
            If b.ValueKind = System.Text.Json.JsonValueKind.Object Then
                Dim tmp As System.Text.Json.JsonElement
                If b.TryGetProperty("text", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then text = tmp.GetString()
                If b.TryGetProperty("level", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then level = tmp.GetInt32()
            ElseIf b.ValueKind = System.Text.Json.JsonValueKind.String Then
                text = b.GetString()
            End If
            If System.String.IsNullOrWhiteSpace(text) Then Continue For
            level = System.Math.Max(0, System.Math.Min(8, level))

            Dim templateP As DocumentFormat.OpenXml.Drawing.Paragraph = templates.FirstOrDefault(Function(p)
                                                                                                     Dim l As System.Int32 = If(p.ParagraphProperties?.Level IsNot Nothing, CInt(p.ParagraphProperties.Level.Value), 0)
                                                                                                     Return l = level
                                                                                                 End Function)
            If templateP Is Nothing Then templateP = templates.FirstOrDefault()

            Dim pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = If(templateP?.ParagraphProperties IsNot Nothing,
                CType(templateP.ParagraphProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.ParagraphProperties),
                New DocumentFormat.OpenXml.Drawing.ParagraphProperties())
            pPr.Level = level
            ' bullet_text explicitly requests a bullet paragraph. Remove a local NoBullet so the layout/master list style can apply.
            pPr.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.NoBullet)()
            ApplyParagraphStyleOverrides(pPr, el)
            ApplyBulletOverride(pPr, el)

            Dim rp As DocumentFormat.OpenXml.Drawing.RunProperties = CloneTemplateRunProperties(templateP)
            ApplyRunStyleOverrides(rp, el)
            Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(text)))
            If templateP IsNot Nothing Then
                Dim templateEndProperties As DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties =
                    templateP.GetFirstChild(Of DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties)()
                If templateEndProperties IsNot Nothing Then
                    para.Append(CType(templateEndProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties))
                End If
            End If
            newTb.Append(para)
            added += 1
        Next

        If added = 0 Then newTb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()))
        targetShape.TextBody = newTb
    End Sub

    Private Sub AppendPreservedTextBodyShell(
        ByVal target As DocumentFormat.OpenXml.Presentation.TextBody,
        ByVal source As DocumentFormat.OpenXml.Presentation.TextBody)

        If source?.BodyProperties IsNot Nothing Then
            target.Append(CType(source.BodyProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.BodyProperties))
        Else
            target.Append(New DocumentFormat.OpenXml.Drawing.BodyProperties())
        End If
        If source?.ListStyle IsNot Nothing Then
            target.Append(CType(source.ListStyle.CloneNode(True), DocumentFormat.OpenXml.Drawing.ListStyle))
        Else
            target.Append(New DocumentFormat.OpenXml.Drawing.ListStyle())
        End If
    End Sub

    Private Function CloneTemplateRunProperties(ByVal templateP As DocumentFormat.OpenXml.Drawing.Paragraph) As DocumentFormat.OpenXml.Drawing.RunProperties
        Dim templateRun As DocumentFormat.OpenXml.Drawing.Run = templateP?.Elements(Of DocumentFormat.OpenXml.Drawing.Run)().FirstOrDefault()
        If templateRun?.RunProperties IsNot Nothing Then Return CType(templateRun.RunProperties.CloneNode(True), DocumentFormat.OpenXml.Drawing.RunProperties)
        Return New DocumentFormat.OpenXml.Drawing.RunProperties()
    End Function

    Private Sub SetRunSolidColorOrdered(
        ByVal rp As DocumentFormat.OpenXml.Drawing.RunProperties,
        ByVal color As System.String)

        If rp Is Nothing Then Return

        ' CT_TextCharacterProperties requires fill properties before the font elements.
        ' Appending a:solidFill after a:latin creates invalid DrawingML and PowerPoint may repair the file.
        For Each child As DocumentFormat.OpenXml.OpenXmlElement In rp.ChildElements.ToList()
            Select Case child.LocalName
                Case "noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill"
                    child.Remove()
            End Select
        Next

        Dim fill As New DocumentFormat.OpenXml.Drawing.SolidFill(
            New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {
                .Val = NormalizeHexColor(color, "#0F172A").TrimStart("#"c)
            })

        Dim insertBefore As DocumentFormat.OpenXml.OpenXmlElement = Nothing
        For Each child As DocumentFormat.OpenXml.OpenXmlElement In rp.ChildElements
            Select Case child.LocalName
                Case "effectLst", "effectDag", "highlight", "latin", "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"
                    insertBefore = child
                    Exit For
            End Select
        Next

        If insertBefore Is Nothing Then
            rp.Append(fill)
        Else
            rp.InsertBefore(fill, insertBefore)
        End If
    End Sub

    Private Sub SetRunLatinFontOrdered(
        ByVal rp As DocumentFormat.OpenXml.Drawing.RunProperties,
        ByVal fontName As System.String)

        If rp Is Nothing OrElse System.String.IsNullOrWhiteSpace(fontName) Then Return
        rp.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.LatinFont)()

        Dim latin As New DocumentFormat.OpenXml.Drawing.LatinFont() With {.Typeface = fontName}
        Dim insertBefore As DocumentFormat.OpenXml.OpenXmlElement = Nothing
        For Each child As DocumentFormat.OpenXml.OpenXmlElement In rp.ChildElements
            Select Case child.LocalName
                Case "ea", "cs", "sym", "hlinkClick", "hlinkMouseOver", "rtl", "extLst"
                    insertBefore = child
                    Exit For
            End Select
        Next

        If insertBefore Is Nothing Then
            rp.Append(latin)
        Else
            rp.InsertBefore(latin, insertBefore)
        End If
    End Sub

    Private Sub ApplyRunStyleOverrides(ByVal rp As DocumentFormat.OpenXml.Drawing.RunProperties, ByVal el As System.Text.Json.JsonElement)
        If rp Is Nothing Then Return
        Dim styleEl As System.Text.Json.JsonElement
        If Not el.TryGetProperty("style", styleEl) OrElse styleEl.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return
        Dim tmp As System.Text.Json.JsonElement

        Dim fontName As System.String = System.String.Empty
        If styleEl.TryGetProperty("fontFamily", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then fontName = tmp.GetString()
        If System.String.IsNullOrWhiteSpace(fontName) AndAlso styleEl.TryGetProperty("font", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then fontName = tmp.GetString()

        If styleEl.TryGetProperty("fontSize", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then rp.FontSize = CInt(tmp.GetDouble() * 100.0R)
        If styleEl.TryGetProperty("size", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then rp.FontSize = CInt(tmp.GetDouble() * 100.0R)
        If styleEl.TryGetProperty("bold", tmp) AndAlso (tmp.ValueKind = System.Text.Json.JsonValueKind.True OrElse tmp.ValueKind = System.Text.Json.JsonValueKind.False) Then rp.Bold = tmp.GetBoolean()
        If styleEl.TryGetProperty("italic", tmp) AndAlso (tmp.ValueKind = System.Text.Json.JsonValueKind.True OrElse tmp.ValueKind = System.Text.Json.JsonValueKind.False) Then rp.Italic = tmp.GetBoolean()

        ' Apply the fill before the Latin font to preserve the schema order in a:rPr.
        If styleEl.TryGetProperty("color", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then
            SetRunSolidColorOrdered(rp, tmp.GetString())
        End If
        If Not System.String.IsNullOrWhiteSpace(fontName) Then
            SetRunLatinFontOrdered(rp, fontName)
        End If
    End Sub

    Private Sub ApplyParagraphStyleOverrides(ByVal pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties, ByVal el As System.Text.Json.JsonElement)
        If pPr Is Nothing Then Return
        Dim styleEl As System.Text.Json.JsonElement
        If Not el.TryGetProperty("style", styleEl) OrElse styleEl.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return
        Dim tmp As System.Text.Json.JsonElement
        If styleEl.TryGetProperty("align", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then
            Select Case tmp.GetString().ToLowerInvariant()
                Case "center" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center
                Case "right" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right
                Case "justify" : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Justified
                Case Else : pPr.Alignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left
            End Select
        End If
    End Sub

    Private Sub ApplyBulletOverride(ByVal pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties, ByVal el As System.Text.Json.JsonElement)
        Dim styleEl As System.Text.Json.JsonElement
        If Not el.TryGetProperty("style", styleEl) OrElse styleEl.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return
        Dim tmp As System.Text.Json.JsonElement
        If styleEl.TryGetProperty("bulletChar", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String AndAlso Not System.String.IsNullOrEmpty(tmp.GetString()) Then
            RemoveBulletDefinition(pPr)
            pPr.Append(New DocumentFormat.OpenXml.Drawing.CharacterBullet() With {.Char = tmp.GetString()(0)})
        End If
    End Sub

    Private Sub RemoveBulletDefinition(ByVal pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties)
        If pPr Is Nothing Then Return
        pPr.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.NoBullet)()
        pPr.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.CharacterBullet)()
        pPr.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.AutoNumberedBullet)()
        pPr.RemoveAllChildren(Of DocumentFormat.OpenXml.Drawing.PictureBullet)()
    End Sub

    Private Function BuildParagraph(
        ByVal text As System.String,
        ByVal el As System.Text.Json.JsonElement,
        Optional ByVal pPr As DocumentFormat.OpenXml.Drawing.ParagraphProperties = Nothing) As DocumentFormat.OpenXml.Drawing.Paragraph

        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
        ApplyRunStyleOverrides(rp, el)
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph()
        If pPr IsNot Nothing Then para.Append(CType(pPr.CloneNode(True), DocumentFormat.OpenXml.Drawing.ParagraphProperties))
        para.Append(New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(If(text, System.String.Empty))))
        Return para
    End Function


    ''' <summary>
    ''' Generates a unique slide key combining sanitized title and slide ID.
    ''' </summary>
    ''' <param name="sp">The slide part to extract title from.</param>
    ''' <param name="slideId">The slide's unique ID.</param>
    ''' <returns>Slide key in format "{SanitizedTitle}-{slideId}" or "SID-{slideId}" if no title.</returns>
    Private Function GetSlideKey(
        sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        slideId As UInteger
      ) As String
        Dim title = GetSlideTitle(sp)
        If String.IsNullOrWhiteSpace(title) Then
            Return $"SID-{slideId}"
        Else
            Return $"{SanitizeKey(title)}-{slideId}"
        End If
    End Function

    ''' <summary>
    ''' Creates or updates speaker notes for a slide with the specified text content.
    ''' </summary>
    ''' <param name="sp">The slide part to add or update notes for.</param>
    ''' <param name="notesText">The notes text content.</param>
    ''' <remarks>
    ''' Creates a new NotesSlidePart if none exists. Removes existing shapes and inserts
    ''' a new body placeholder at index 2 (after header elements).
    ''' </remarks>
    Private Sub SetSpeakerNotes(
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    notesText As String)

        Dim notesPart As DocumentFormat.OpenXml.Packaging.NotesSlidePart = sp.NotesSlidePart
        If notesPart Is Nothing Then
            notesPart = sp.AddNewPart(Of DocumentFormat.OpenXml.Packaging.NotesSlidePart)()
            notesPart.NotesSlide = New DocumentFormat.OpenXml.Presentation.NotesSlide(
            New DocumentFormat.OpenXml.Presentation.CommonSlideData(
                New DocumentFormat.OpenXml.Presentation.ShapeTree(
                    New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeProperties(
                        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 1UI, .Name = ""},
                        New DocumentFormat.OpenXml.Presentation.NonVisualGroupShapeDrawingProperties(),
                        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()),
                    New DocumentFormat.OpenXml.Presentation.GroupShapeProperties())),
            New DocumentFormat.OpenXml.Presentation.ColorMapOverride(
                New DocumentFormat.OpenXml.Drawing.MasterColorMapping()))
        End If

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree =
        notesPart.NotesSlide.CommonSlideData.ShapeTree

        ' ----- only remove Shapes/Pics -----
        For Each n In tree.ChildElements.OfType(Of DocumentFormat.OpenXml.OpenXmlElement)().ToList()
            If TypeOf n Is DocumentFormat.OpenXml.Presentation.Shape _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.Picture _
           OrElse TypeOf n Is DocumentFormat.OpenXml.Presentation.GroupShape Then
                n.Remove()
            End If
        Next

        ' ----- new Body-Shape -----
        Dim nvSpPr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
        New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = 2UI, .Name = "NotesBody"},
        New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(
            New DocumentFormat.OpenXml.Drawing.ShapeLocks() With {.NoGrouping = True}),
        New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties(
            New DocumentFormat.OpenXml.Presentation.PlaceholderShape() With {
                .Type = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                .Index = 1UI}))
        Dim shapePr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        Dim noteShape As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, shapePr)

        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(
        New DocumentFormat.OpenXml.Drawing.BodyProperties(),
        New DocumentFormat.OpenXml.Drawing.ListStyle())
        Dim run As New DocumentFormat.OpenXml.Drawing.Run(
        New DocumentFormat.OpenXml.Drawing.RunProperties(),
        New DocumentFormat.OpenXml.Drawing.Text(notesText))
        Dim para As New DocumentFormat.OpenXml.Drawing.Paragraph(run) With {
        .ParagraphProperties = New DocumentFormat.OpenXml.Drawing.ParagraphProperties()}
        tb.Append(para)
        noteShape.Append(tb)

        ' insert after header
        If tree.ChildElements.Count >= 2 Then
            tree.InsertAt(noteShape, 2)
        Else
            tree.Append(noteShape)
        End If

        notesPart.NotesSlide.Save()
    End Sub



    ''' <summary>
    ''' Creates an OpenXML Fill element from a JSON definition.
    ''' </summary>
    ''' <param name="fillJson">JSON element containing fill type and color properties.</param>
    ''' <returns>SolidFill element with specified color, or NoFill as fallback.</returns>

    Private Function CreateFill(fillJson As JsonElement) As DocumentFormat.OpenXml.OpenXmlElement
        Dim fillType As String = ""
        If fillJson.TryGetProperty("type", Nothing) Then fillType = fillJson.GetProperty("type").GetString()

        Select Case fillType.ToLower()
            Case "solid"
                If fillJson.TryGetProperty("color", Nothing) Then
                    Dim colorHex = fillJson.GetProperty("color").GetString().TrimStart("#"c)
                    Return New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorHex})
                End If
        End Select
        Return New DocumentFormat.OpenXml.Drawing.NoFill() ' Fallback
    End Function

    ''' <summary>
    ''' Creates an OpenXML Outline element from a JSON definition.
    ''' </summary>
    ''' <param name="outlineJson">JSON element containing width, color, and dashType properties.</param>
    ''' <returns>Configured Outline element with width in EMUs, color, and dash style.</returns>
    ''' <remarks>Width is parsed culture-invariantly; 1 point = 12700 EMUs.</remarks>
    Private Function CreateOutline(outlineJson As JsonElement) As DocumentFormat.OpenXml.Drawing.Outline
        Dim outline As New DocumentFormat.OpenXml.Drawing.Outline()

        Dim widthJson As JsonElement
        If outlineJson.TryGetProperty("width", widthJson) Then
            ' [FIX] This safely parses numbers like "1" or "1.5" regardless of system language.
            Dim widthValue As Double
            If Double.TryParse(widthJson.GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, widthValue) Then
                outline.Width = CInt(widthValue * 12700) ' 1 point = 12700 EMUs
            End If
        End If

        Dim colorJson As JsonElement
        If outlineJson.TryGetProperty("color", colorJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex With {.Val = colorJson.GetString().TrimStart("#"c)}))
        End If

        Dim dashJson As JsonElement
        If outlineJson.TryGetProperty("dashType", dashJson) Then
            outline.Append(New DocumentFormat.OpenXml.Drawing.PresetDash With {.Val = JsonDashNameToEnumValue(dashJson.GetString())})
        End If

        Return outline
    End Function



    ''' <summary>
    ''' Converts relative percentage-based coordinates from JSON into absolute EMU coordinates.
    ''' </summary>
    ''' <param name="presPart">The presentation part, to get the master slide dimensions.</param>
    ''' <param name="transformJson">The JSON "transform" object.</param>
    ''' <returns>A fully calculated Transform2D object with absolute EMUs.</returns>
    Private Function ConvertRelativeToAbsoluteTransform(presPart As DocumentFormat.OpenXml.Packaging.PresentationPart, transformJson As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Drawing.Transform2D
        ' Get the master slide dimensions in EMUs
        Dim slideWidthEmu = presPart.Presentation.SlideSize.Cx.Value
        Dim slideHeightEmu = presPart.Presentation.SlideSize.Cy.Value

        ' Safely parse the relative percentage values from JSON
        Dim relX, relY, relW, relH As Double
        Double.TryParse(transformJson.GetProperty("x").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relX)
        Double.TryParse(transformJson.GetProperty("y").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relY)
        Double.TryParse(transformJson.GetProperty("width").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relW)
        Double.TryParse(transformJson.GetProperty("height").GetRawText(), Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, relH)

        ' Calculate the absolute EMU values
        Dim absX = CLng(slideWidthEmu * relX)
        Dim absY = CLng(slideHeightEmu * relY)
        Dim absCx = CLng(slideWidthEmu * relW)
        Dim absCy = CLng(slideHeightEmu * relH)

        Return New DocumentFormat.OpenXml.Drawing.Transform2D(
        New DocumentFormat.OpenXml.Drawing.Offset With {.X = absX, .Y = absY},
        New DocumentFormat.OpenXml.Drawing.Extents With {.Cx = absCx, .Cy = absCy}
    )
    End Function


    ''' <summary>
    ''' Converts a JSON shape name string to the corresponding OpenXML ShapeTypeValues enum.
    ''' </summary>
    ''' <param name="jsonName">Case-insensitive shape name (e.g., "rectangle", "circle").</param>
    ''' <returns>Matching ShapeTypeValues enum; defaults to Rectangle if unknown.</returns>
    Private Function JsonShapeNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.ShapeTypeValues
        Select Case jsonName.ToLower()
            Case "rectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
            Case "oval", "ellipse", "circle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse
            Case "line" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line
            Case "rightarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RightArrow
            Case "leftarrow" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.LeftArrow
            Case "triangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Triangle ' Corrected from IsoscelesTriangle
            Case "roundedrectangle" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle
            Case "flowchartprocess" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartProcess
            Case "flowchartdecision" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartDecision
            Case "flowchartterminator" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.FlowChartTerminator
            Case "chevron" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Chevron
            Case "pentagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Pentagon
            Case "hexagon" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Hexagon
            Case "plus" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Plus
            Case "blockarc" : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.BlockArc
            Case Else : Return DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle ' Fallback
        End Select
    End Function


    ''' <summary>
    ''' Converts a JSON dash style name to the corresponding OpenXML PresetLineDashValues enum.
    ''' </summary>
    ''' <param name="jsonName">Case-insensitive dash name (e.g., "solid", "dashed", "dotted").</param>
    ''' <returns>Matching PresetLineDashValues enum; defaults to Solid if unknown.</returns>
    Private Function JsonDashNameToEnumValue(jsonName As String) As DocumentFormat.OpenXml.Drawing.PresetLineDashValues
        Select Case jsonName.ToLower()
            Case "solid"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid
            Case "dot", "dotted"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dot
            Case "dash", "dashed"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Dash
            Case "longdash"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDash
            Case "dashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.DashDot
            Case "longdashdot"
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.LargeDashDot
            Case Else
                Return DocumentFormat.OpenXml.Drawing.PresetLineDashValues.Solid ' Fallback
        End Select
    End Function


    ''' <summary>
    ''' Builds a Drawing.Paragraph with text, nesting level, and optional bullet formatting.
    ''' </summary>
    ''' <param name="text">The text content.</param>
    ''' <param name="level">Nesting level for bullet indentation.</param>
    ''' <param name="el">JSON element containing style properties (font, size, bold, color, align).</param>
    ''' <param name="isBulleted">True to use default bullets; False to apply NoBullet.</param>
    ''' <returns>Configured Paragraph element with style and bullet properties.</returns>
    Private Function BuildStyledParagraph(
        ByVal text As System.String,
        ByVal level As System.Int32,
        ByVal el As System.Text.Json.JsonElement,
        ByVal isBulleted As System.Boolean) As DocumentFormat.OpenXml.Drawing.Paragraph

        Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {.Level = System.Math.Max(0, System.Math.Min(8, level))}
        If Not isBulleted Then pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        ApplyParagraphStyleOverrides(pPr, el)
        If isBulleted Then ApplyBulletOverride(pPr, el)
        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
        ApplyRunStyleOverrides(rp, el)
        Return New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(If(text, System.String.Empty))))
    End Function

    Private Sub CreateFreestandingTextBox(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree
        Dim id As UInteger = NextShapeId(sp)
        Dim tf As System.Text.Json.JsonElement
        If Not el.TryGetProperty("transform", tf) Then Return
        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetTransformFromJson(presPart, tf)

        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = xfrm}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle})
        spPr.Append(New DocumentFormat.OpenXml.Drawing.NoFill())
        spPr.Append(New DocumentFormat.OpenXml.Drawing.Outline(New DocumentFormat.OpenXml.Drawing.NoFill()))

        Dim nvDr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties() With {.TextBox = True}
        Dim nv As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = id, .Name = "TextBox " & id.ToString(System.Globalization.CultureInfo.InvariantCulture)},
            nvDr,
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(New DocumentFormat.OpenXml.Drawing.BodyProperties(), New DocumentFormat.OpenXml.Drawing.ListStyle())
        Dim typeName As System.String = JsonString(el, "type", "text").ToLowerInvariant()
        If typeName = "bullet_text" Then
            Dim bullets As System.Text.Json.JsonElement
            If el.TryGetProperty("bullets", bullets) AndAlso bullets.ValueKind = System.Text.Json.JsonValueKind.Array Then
                For Each b As System.Text.Json.JsonElement In bullets.EnumerateArray()
                    Dim txt As System.String = If(b.ValueKind = System.Text.Json.JsonValueKind.Object, JsonString(b, "text", ""), If(b.ValueKind = System.Text.Json.JsonValueKind.String, b.GetString(), ""))
                    If System.String.IsNullOrWhiteSpace(txt) Then Continue For
                    Dim lvl As System.Int32 = If(b.ValueKind = System.Text.Json.JsonValueKind.Object, CInt(JsonDouble(b, "level", 0.0R)), 0)
                    Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {
                        .Level = System.Math.Max(0, System.Math.Min(8, lvl)),
                        .LeftMargin = 360000,
                        .Indent = -180000
                    }
                    ApplyParagraphStyleOverrides(pPr, el)
                    ApplyBulletOverride(pPr, el)
                    If Not pPr.ChildElements.Any(Function(c) TypeOf c Is DocumentFormat.OpenXml.Drawing.CharacterBullet OrElse TypeOf c Is DocumentFormat.OpenXml.Drawing.AutoNumberedBullet OrElse TypeOf c Is DocumentFormat.OpenXml.Drawing.PictureBullet) Then
                        pPr.Append(New DocumentFormat.OpenXml.Drawing.CharacterBullet() With {.Char = "•"c})
                    End If
                    Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties()
                    ApplyRunStyleOverrides(rp, el)
                    tb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(txt))))
                Next
            End If
        Else
            tb.Append(BuildParagraph(JsonString(el, "text", ""), el))
        End If
        If Not tb.Elements(Of DocumentFormat.OpenXml.Drawing.Paragraph)().Any() Then tb.Append(New DocumentFormat.OpenXml.Drawing.Paragraph(New DocumentFormat.OpenXml.Drawing.EndParagraphRunProperties()))

        tree.Append(New DocumentFormat.OpenXml.Presentation.Shape(nv, spPr, tb))
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Adds a geometric shape to a slide with optional fill, outline, and text content.
    ''' </summary>
    ''' <param name="presPart">The presentation part for slide dimensions.</param>
    ''' <param name="sp">The slide part to add the shape to.</param>
    ''' <param name="el">JSON element containing shapeType, transform, fill, outline, and optional text properties.</param>
    ''' <remarks>
    ''' Transform coordinates can be in EMUs (>1) or percentages (≤1).
    ''' Supports various shape types via JsonShapeNameToEnumValue mapping.
    ''' </remarks>
    Private Sub AddShape(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree As DocumentFormat.OpenXml.Presentation.ShapeTree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) Determine next available shape ID
        Dim maxId As UInteger = 0UI
        For Each nvPr As DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties _
        In tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)()
            If nvPr.Id.Value > maxId Then maxId = nvPr.Id.Value
        Next
        Dim newId As UInteger = maxId + 1UI

        ' 2) Extract transform from JSON
        Dim transformJson = el.GetProperty("transform")

        ' Check raw value: ≤1 = percentage, >1 = already EMUs
        Dim rawX As Double
        If Not Double.TryParse(transformJson.GetProperty("x").GetRawText(),
                       Globalization.NumberStyles.Any,
                       Globalization.CultureInfo.InvariantCulture,
                       rawX) Then
            rawX = 0.0
        End If

        Dim absoluteTransform As DocumentFormat.OpenXml.Drawing.Transform2D

        If rawX <= 1.0 Then
            ' Percentage values → convert to EMU
            absoluteTransform = ConvertRelativeToAbsoluteTransform(presPart, transformJson)
        Else
            ' Use direct EMU values
            Dim ofs As New DocumentFormat.OpenXml.Drawing.Offset() With {
            .X = CLng(transformJson.GetProperty("x").GetInt64()),
        .Y = CLng(transformJson.GetProperty("y").GetInt64())
    }
            Dim ext As New DocumentFormat.OpenXml.Drawing.Extents() With {
        .Cx = CLng(transformJson.GetProperty("width").GetInt64()),
        .Cy = CLng(transformJson.GetProperty("height").GetInt64())
    }
            absoluteTransform = New DocumentFormat.OpenXml.Drawing.Transform2D(ofs, ext)
        End If

        ' 3) ShapeProperties
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = absoluteTransform}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = JsonShapeNameToEnumValue(el.GetProperty("shapeType").GetString())})
        If el.TryGetProperty("fill", Nothing) Then spPr.Append(CreateFill(el.GetProperty("fill")))
        If el.TryGetProperty("outline", Nothing) Then spPr.Append(CreateOutline(el.GetProperty("outline")))

        ' 4) nvSpPr (only set TextBox if text content follows)
        Dim nvSpDr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties()
        If el.TryGetProperty("text", Nothing) Then
            nvSpDr.TextBox = True
            nvSpDr.AppendChild(New DocumentFormat.OpenXml.Drawing.ShapeLocks())
        End If
        Dim nvSpPr = New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = newId, .Name = $"Shape {newId}"},
    nvSpDr,
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties()
)
        ' 5) Create new Shape
        Dim shp As New DocumentFormat.OpenXml.Presentation.Shape(nvSpPr, spPr)

        ' 6) Optional text content
        If el.TryGetProperty("text", Nothing) Then
            Dim tb = New DocumentFormat.OpenXml.Presentation.TextBody(
            New DocumentFormat.OpenXml.Drawing.BodyProperties(),
            New DocumentFormat.OpenXml.Drawing.ListStyle()
        )
            tb.Append(BuildStyledParagraph(el.GetProperty("text").GetString(), 0, el, False))
            shp.Append(tb)
        End If

        ' 7) Insert and save
        tree.Append(shp)
        sp.Slide.Save()
    End Sub

    ''' <summary>
    ''' Inserts an SVG icon from the JSON at the given location.
    ''' Uses a standard <p:pic> with <a:blip>; this is the same recipe
    ''' PowerPoint 2019+ generates and shows on Office 2016 (Oct-2018) too.
    ''' </summary>
    Private Sub AddSvgIcon(
    presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
    sp As DocumentFormat.OpenXml.Packaging.SlidePart,
    el As System.Text.Json.JsonElement)

        Dim tree = sp.Slide.CommonSlideData.ShapeTree

        ' 1) unique ID on slide
        Dim newId As UInteger =
    tree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)().
        Select(Function(nv) nv.Id.Value).DefaultIfEmpty(0).Max() + 1UI

        ' 2) build Transform2D (percent → EMU if needed)
        Dim tf = el.GetProperty("transform")
        Dim rawX As Double
        Double.TryParse(tf.GetProperty("x").GetRawText(),
        Globalization.NumberStyles.Any,
        Globalization.CultureInfo.InvariantCulture,
        rawX)

        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D
        If rawX > 1 Then
            xfrm = New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset With {
                .X = CLng(tf.GetProperty("x").GetInt64()),
                .Y = CLng(tf.GetProperty("y").GetInt64())},
            New DocumentFormat.OpenXml.Drawing.Extents With {
                .Cx = CLng(tf.GetProperty("width").GetInt64()),
                .Cy = CLng(tf.GetProperty("height").GetInt64())})
        Else
            ' Assuming ConvertRelativeToAbsoluteTransform exists and returns a Transform2D
            ' xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf) 
            xfrm = ConvertRelativeToAbsoluteTransform(presPart, tf)
        End If

        ' 3) embed SVG file
        Dim svgPart = sp.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Svg)
        Using ms As New IO.MemoryStream(
        System.Text.Encoding.UTF8.GetBytes(el.GetProperty("svg").GetString()))
            svgPart.FeedData(ms)
        End Using
        Dim relId As String = sp.GetIdOfPart(svgPart)

        ' 4) build <p:pic>
        Dim nvPic As New DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties(
    New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {
        .Id = newId, .Name = "Icon " & newId},
    New DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties(
        New DocumentFormat.OpenXml.Drawing.PictureLocks() With {.NoChangeAspect = True}),
    New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim blipFill As New DocumentFormat.OpenXml.Presentation.BlipFill(
    New DocumentFormat.OpenXml.Drawing.Blip() With {
        .Embed = relId,
        .CompressionState =
            DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print},
    New DocumentFormat.OpenXml.Drawing.Stretch(
        New DocumentFormat.OpenXml.Drawing.FillRectangle()))

        ' Define the rectangle geometry
        Dim prstGeom As New DocumentFormat.OpenXml.Drawing.PresetGeometry(
        New DocumentFormat.OpenXml.Drawing.AdjustValueList()
    ) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle}

        ' Create the ShapeProperties object
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()

        ' Append the transform and geometry as child elements
        spPr.Append(xfrm)
        spPr.Append(prstGeom)

        ' Create the final picture by combining all the parts
        Dim pic As New DocumentFormat.OpenXml.Presentation.Picture(nvPic, blipFill, spPr)

        ' 5) append & save
        tree.Append(pic)
        sp.Slide.Save()
    End Sub


    ''' <summary>
    ''' Removes empty body placeholder shapes from a slide.
    ''' </summary>
    ''' <param name="sp">The slide part to process.</param>
    ''' <remarks>
    ''' A body placeholder is considered empty if it has no TextBody or contains only whitespace.
    ''' Only removes the first empty body placeholder found.
    ''' </remarks>
    Private Sub RemoveEmptyBodyPlaceholder(sp As DocumentFormat.OpenXml.Packaging.SlidePart)
        Dim shpToRemove As DocumentFormat.OpenXml.Presentation.Shape = Nothing

        For Each shp In sp.Slide.CommonSlideData.ShapeTree.
                         Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()

            Dim ph = shp.NonVisualShapeProperties?.
                        ApplicationNonVisualDrawingProperties?.
                        PlaceholderShape
            If ph IsNot Nothing AndAlso ph.Type IsNot Nothing AndAlso
               ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body Then

                Dim semanticRole As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shp)

                ' Never remove corporate artwork that happens to be implemented as a Body
                ' placeholder. The VISCHER closing logo is exactly such a shape.
                If Not System.String.Equals(semanticRole, "body", System.StringComparison.OrdinalIgnoreCase) OrElse
                   ShapeContainsRelationshipBackedVisual(shp) Then
                    Continue For
                End If

                ' empty = only one paragraph with no text or whitespace
                Dim empty As Boolean =
                    (shp.TextBody Is Nothing) OrElse
                    Not shp.TextBody.Descendants(Of DocumentFormat.OpenXml.Drawing.Text)().
                        Any(Function(t) Not String.IsNullOrWhiteSpace(t.Text))

                If empty Then shpToRemove = shp
                Exit For
            End If
        Next

        If shpToRemove IsNot Nothing Then
            shpToRemove.Remove()
            sp.Slide.Save()
        End If
    End Sub

    ''' <summary>
    ''' Determines whether a placeholder represents a body-like content area.
    ''' </summary>
    ''' <param name="ph">The placeholder shape to check.</param>
    ''' <returns>True if the placeholder is Body or Object type; False otherwise.</returns>
    ''' <remarks>
    ''' Excludes footer, date/time, slide number, title, and subtitle placeholders.
    ''' Implicit placeholders (no type) are treated as body-like only if index >= 2.
    ''' </remarks>
    Private Function IsBodyLikePlaceholder(ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape) As Boolean
        If ph Is Nothing Then Return False
        If ph.Type Is Nothing Then
            ' Implicit placeholder: treat as body-like only if index is typical for content (not header/footer indices)
            ' Common patterns: Title = index 0, Subtitle = index 1; Footer/Date/SlideNumber often have explicit types.
            ' With no type, be conservative: don't auto-accept implicit unless index >= 2.
            If ph.Index IsNot Nothing Then
                Return ph.Index.Value >= 2UI
            End If
            Return False
        End If

        Select Case ph.Type.Value
            Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Body,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object
                Return True
            Case DocumentFormat.OpenXml.Presentation.PlaceholderValues.Footer,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.DateAndTime,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.SlideNumber,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle,
                 DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle
                Return False
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Sets plain text in a shape identified by placeholder element.
    ''' </summary>
    ''' <param name="sp">The slide part containing the shape.</param>
    ''' <param name="placeholderEl">JSON element identifying the placeholder.</param>
    ''' <param name="text">The text to insert.</param>
    ''' <param name="el">JSON element containing style properties.</param>

#Region "Professional Visual Components"

    Private Structure EmuRect
        Public X As Long
        Public Y As Long
        Public Width As Long
        Public Height As Long
    End Structure

    Private NotInheritable Class RenderStyleContext
        Public Property Dark As System.String
        Public Property Light As System.String
        Public Property Surface As System.String
        Public Property MutedSurface As System.String
        Public Property MutedText As System.String
        Public Property Accent1 As System.String
        Public Property Accent2 As System.String
        Public Property Accent3 As System.String
        Public Property HeadingFont As System.String
        Public Property BodyFont As System.String
        Public Property HeadingSize As System.Double
        Public Property BodySize As System.Double
        Public Property SmallSize As System.Double
        Public Property Background As System.String
        Public Property IsDarkBackground As System.Boolean
    End Class

    Private Function NormalizeHexColor(ByVal value As System.String, ByVal fallback As System.String) As System.String
        Dim candidate As System.String = If(value, System.String.Empty).Trim().TrimStart("#"c)
        If candidate.Length = 3 Then candidate = System.String.Concat(candidate.Select(Function(c) New System.String(c, 2)))
        If candidate.Length <> 6 OrElse Not candidate.All(Function(c) System.Uri.IsHexDigit(c)) Then Return If(fallback.StartsWith("#", System.StringComparison.Ordinal), fallback.ToUpperInvariant(), "#" & fallback.ToUpperInvariant())
        Return "#" & candidate.ToUpperInvariant()
    End Function

    Private Function MixHexColors(ByVal first As System.String, ByVal second As System.String, ByVal secondWeight As System.Double) As System.String
        first = NormalizeHexColor(first, "#0F172A")
        second = NormalizeHexColor(second, "#FFFFFF")
        secondWeight = System.Math.Max(0.0R, System.Math.Min(1.0R, secondWeight))
        Dim r1 As System.Int32 = System.Convert.ToInt32(first.Substring(1, 2), 16)
        Dim g1 As System.Int32 = System.Convert.ToInt32(first.Substring(3, 2), 16)
        Dim b1 As System.Int32 = System.Convert.ToInt32(first.Substring(5, 2), 16)
        Dim r2 As System.Int32 = System.Convert.ToInt32(second.Substring(1, 2), 16)
        Dim g2 As System.Int32 = System.Convert.ToInt32(second.Substring(3, 2), 16)
        Dim b2 As System.Int32 = System.Convert.ToInt32(second.Substring(5, 2), 16)
        Dim r As System.Int32 = CInt(System.Math.Round(r1 * (1.0R - secondWeight) + r2 * secondWeight))
        Dim g As System.Int32 = CInt(System.Math.Round(g1 * (1.0R - secondWeight) + g2 * secondWeight))
        Dim b As System.Int32 = CInt(System.Math.Round(b1 * (1.0R - secondWeight) + b2 * secondWeight))
        Return "#" & r.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) & g.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) & b.ToString("X2", System.Globalization.CultureInfo.InvariantCulture)
    End Function

    Private Function GetRenderStyleContext(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, Optional ByVal focusRect As System.Nullable(Of EmuRect) = Nothing) As RenderStyleContext
        Dim profile As New DesignProfileJson With {
            .Palette = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        }
        Dim theme As DocumentFormat.OpenXml.Drawing.Theme = sp?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme
        FillThemeProfile(theme, profile)

        Dim themeDark As System.String = If(profile.Palette.ContainsKey("dark1"), NormalizeHexColor(profile.Palette("dark1"), "#0F172A"), "#0F172A")
        Dim themeLight As System.String = If(profile.Palette.ContainsKey("light1"), NormalizeHexColor(profile.Palette("light1"), "#F8FAFC"), "#F8FAFC")
        Dim backgroundRef As System.String = ExtractLocalBackgroundColor(sp, focusRect)
        Dim background As System.String = ResolveThemeColorReferenceForSlide(sp, backgroundRef, themeLight)
        Dim darkBackground As System.Boolean = IsDarkHexColor(background)

        Dim ctx As New RenderStyleContext With {
            .Dark = If(darkBackground, themeLight, themeDark),
            .Light = background,
            .Background = background,
            .IsDarkBackground = darkBackground,
            .Accent1 = If(profile.Palette.ContainsKey("accent1"), NormalizeHexColor(profile.Palette("accent1"), "#2563EB"), "#2563EB"),
            .Accent2 = If(profile.Palette.ContainsKey("accent2"), NormalizeHexColor(profile.Palette("accent2"), "#0F766E"), "#0F766E"),
            .Accent3 = If(profile.Palette.ContainsKey("accent3"), NormalizeHexColor(profile.Palette("accent3"), "#D97706"), "#D97706"),
            .HeadingFont = If(System.String.IsNullOrWhiteSpace(profile.HeadingFont), "Aptos Display", profile.HeadingFont),
            .BodyFont = If(System.String.IsNullOrWhiteSpace(profile.BodyFont), "Aptos", profile.BodyFont),
            .HeadingSize = 15.0R,
            .BodySize = 12.0R,
            .SmallSize = 10.0R
        }

        ' Prefer the effective placeholder styles of the selected layout/slide over the generic theme.
        Dim titleShape As DocumentFormat.OpenXml.Presentation.Shape = FindFirstShapeByRole(sp, "title")
        Dim bodyShape As DocumentFormat.OpenXml.Presentation.Shape = FindFirstShapeByRole(sp, "body")
        If titleShape IsNot Nothing Then
            Dim style As TextStyleJson = ExtractEffectiveTextStyleForShape(sp, titleShape)
            ApplyEffectiveStyleToRenderContext(sp, ctx, style, theme, isHeading:=True)
        End If
        If bodyShape IsNot Nothing Then
            Dim style As TextStyleJson = ExtractEffectiveTextStyleForShape(sp, bodyShape)
            ApplyEffectiveStyleToRenderContext(sp, ctx, style, theme, isHeading:=False)
        End If

        If IsConsultingGeneratedBackground(ctx.Background) Then
            ctx.Accent1 = "#1F6F8B"
            ctx.Accent2 = "#2A9D8F"
            ctx.Accent3 = "#D97745"
            ctx.Dark = If(ctx.IsDarkBackground, "#F7FAFC", "#142536")
            ctx.BodySize = System.Math.Max(12.5R, System.Math.Min(14.0R, ctx.BodySize))
            ctx.SmallSize = System.Math.Max(10.5R, System.Math.Min(11.5R, ctx.BodySize * 0.86R))
            ctx.HeadingSize = System.Math.Max(14.5R, System.Math.Min(17.0R, ctx.HeadingSize))
        End If

        ctx.Surface = MixHexColors(ctx.Background, ctx.Dark, If(ctx.IsDarkBackground, 0.07R, 0.025R))
        ctx.MutedSurface = MixHexColors(ctx.Accent1, ctx.Background, If(ctx.IsDarkBackground, 0.76R, 0.88R))
        ctx.MutedText = MixHexColors(ctx.Dark, ctx.Background, If(ctx.IsDarkBackground, 0.32R, 0.38R))
        Return ctx
    End Function

    Private Function FindFirstShapeByRole(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal desiredRole As System.String) As DocumentFormat.OpenXml.Presentation.Shape

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return Nothing
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For
            Dim role As System.String = ResolvePlaceholderRoleForJson(ph, If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty))
            If System.String.Equals(role, desiredRole, System.StringComparison.OrdinalIgnoreCase) Then Return shp
        Next
        Return Nothing
    End Function

    Private Sub ApplyEffectiveStyleToRenderContext(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal ctx As RenderStyleContext,
        ByVal style As TextStyleJson,
        ByVal theme As DocumentFormat.OpenXml.Drawing.Theme,
        ByVal isHeading As System.Boolean)

        If ctx Is Nothing OrElse style Is Nothing Then Return
        If Not System.String.IsNullOrWhiteSpace(style.FontFamily) Then
            If isHeading Then
                ctx.HeadingFont = style.FontFamily
            Else
                ctx.BodyFont = style.FontFamily
            End If
        End If

        If style.FontSize.HasValue AndAlso style.FontSize.Value > 0.0R Then
            If isHeading Then
                ' Component headings should be subordinate to the slide title.
                ctx.HeadingSize = System.Math.Max(13.0R, System.Math.Min(18.0R, style.FontSize.Value * 0.58R))
            Else
                ctx.BodySize = System.Math.Max(10.5R, System.Math.Min(15.0R, style.FontSize.Value))
                ctx.SmallSize = System.Math.Max(8.5R, System.Math.Min(12.0R, ctx.BodySize * 0.84R))
            End If
        End If

        If Not isHeading AndAlso Not System.String.IsNullOrWhiteSpace(style.Color) Then
            Dim resolved As System.String = ResolveThemeColorReferenceForSlide(sp, style.Color, ctx.Dark)
            If resolved.StartsWith("#", System.StringComparison.Ordinal) AndAlso IsDarkHexColor(resolved) <> ctx.IsDarkBackground Then
                ctx.Dark = resolved
            End If
        End If
    End Sub

    Private Function ResolveThemeColorReference(
        ByVal theme As DocumentFormat.OpenXml.Drawing.Theme,
        ByVal value As System.String,
        ByVal fallback As System.String) As System.String

        If System.String.IsNullOrWhiteSpace(value) Then Return NormalizeHexColor(fallback, "#0F172A")
        If value.Trim().StartsWith("#", System.StringComparison.Ordinal) Then Return NormalizeHexColor(value, fallback)
        If Not value.StartsWith("scheme:", System.StringComparison.OrdinalIgnoreCase) Then Return NormalizeHexColor(fallback, "#0F172A")

        Dim key As System.String = value.Substring("scheme:".Length).Trim().ToLowerInvariant()
        Dim profile As New DesignProfileJson With {
            .Palette = New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.OrdinalIgnoreCase)
        }
        FillThemeProfile(theme, profile)

        Dim paletteKey As System.String = System.String.Empty
        Select Case key
            Case "text1", "tx1", "dark1", "dk1" : paletteKey = "dark1"
            Case "background1", "bg1", "light1", "lt1" : paletteKey = "light1"
            Case "text2", "tx2", "dark2", "dk2" : paletteKey = "dark2"
            Case "background2", "bg2", "light2", "lt2" : paletteKey = "light2"
            Case "accent1" : paletteKey = "accent1"
            Case "accent2" : paletteKey = "accent2"
            Case "accent3" : paletteKey = "accent3"
            Case "accent4" : paletteKey = "accent4"
            Case "accent5" : paletteKey = "accent5"
            Case "accent6" : paletteKey = "accent6"
        End Select

        If Not System.String.IsNullOrWhiteSpace(paletteKey) AndAlso profile.Palette.ContainsKey(paletteKey) Then
            Return NormalizeHexColor(profile.Palette(paletteKey), fallback)
        End If
        Return NormalizeHexColor(fallback, "#0F172A")
    End Function

    Private Function NormalizeColorMapKey(ByVal value As System.String) As System.String
        If System.String.IsNullOrWhiteSpace(value) Then Return System.String.Empty
        Select Case value.Trim().ToLowerInvariant()
            Case "background1" : Return "bg1"
            Case "text1" : Return "tx1"
            Case "background2" : Return "bg2"
            Case "text2" : Return "tx2"
            Case "dark1" : Return "dk1"
            Case "light1" : Return "lt1"
            Case "dark2" : Return "dk2"
            Case "light2" : Return "lt2"
            Case Else : Return value.Trim().ToLowerInvariant()
        End Select
    End Function

    Private Function GetColorMapAttribute(ByVal mapping As DocumentFormat.OpenXml.OpenXmlElement, ByVal key As System.String) As System.String
        If mapping Is Nothing OrElse System.String.IsNullOrWhiteSpace(key) Then Return System.String.Empty
        For Each attr As DocumentFormat.OpenXml.OpenXmlAttribute In mapping.GetAttributes()
            If System.String.Equals(attr.LocalName, key, System.StringComparison.OrdinalIgnoreCase) Then Return If(attr.Value, System.String.Empty)
        Next
        Return System.String.Empty
    End Function

    Private Function FindOverrideColorMapping(ByVal colorMapOverride As DocumentFormat.OpenXml.OpenXmlElement) As DocumentFormat.OpenXml.OpenXmlElement
        If colorMapOverride Is Nothing Then Return Nothing
        Return colorMapOverride.ChildElements.FirstOrDefault(Function(e) System.String.Equals(e.LocalName, "overrideClrMapping", System.StringComparison.OrdinalIgnoreCase))
    End Function

    Private Function ResolveMappedSchemeKeyForSlide(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal schemeKey As System.String) As System.String

        Dim key As System.String = NormalizeColorMapKey(schemeKey)
        If System.String.IsNullOrWhiteSpace(key) Then Return key
        ' dk/lt are physical theme slots and are not remapped through clrMap.
        If key = "dk1" OrElse key = "lt1" OrElse key = "dk2" OrElse key = "lt2" Then Return key

        Dim mapping As DocumentFormat.OpenXml.OpenXmlElement = FindOverrideColorMapping(sp?.Slide?.ColorMapOverride)
        Dim mapped As System.String = GetColorMapAttribute(mapping, key)
        If Not System.String.IsNullOrWhiteSpace(mapped) Then Return NormalizeColorMapKey(mapped)

        mapping = FindOverrideColorMapping(sp?.SlideLayoutPart?.SlideLayout?.ColorMapOverride)
        mapped = GetColorMapAttribute(mapping, key)
        If Not System.String.IsNullOrWhiteSpace(mapped) Then Return NormalizeColorMapKey(mapped)

        Dim masterMap As DocumentFormat.OpenXml.OpenXmlElement = sp?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.ColorMap
        mapped = GetColorMapAttribute(masterMap, key)
        If Not System.String.IsNullOrWhiteSpace(mapped) Then Return NormalizeColorMapKey(mapped)
        Return key
    End Function

    Private Function ResolveThemeColorReferenceForSlide(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal value As System.String,
        ByVal fallback As System.String) As System.String

        If System.String.IsNullOrWhiteSpace(value) OrElse Not value.StartsWith("scheme:", System.StringComparison.OrdinalIgnoreCase) Then
            Return ResolveThemeColorReference(sp?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme, value, fallback)
        End If
        Dim mapped As System.String = ResolveMappedSchemeKeyForSlide(sp, value.Substring("scheme:".Length))
        Return ResolveThemeColorReference(sp?.SlideLayoutPart?.SlideMasterPart?.ThemePart?.Theme, "scheme:" & mapped, fallback)
    End Function

    Private Function RectContainsCenter(ByVal candidate As EmuRect, ByVal focus As EmuRect) As System.Boolean
        Dim cx As System.Double = focus.X + focus.Width / 2.0R
        Dim cy As System.Double = focus.Y + focus.Height / 2.0R
        Return cx >= candidate.X AndAlso cx <= candidate.X + candidate.Width AndAlso cy >= candidate.Y AndAlso cy <= candidate.Y + candidate.Height
    End Function

    Private Sub ConsiderBackgroundShapes(
        ByVal tree As DocumentFormat.OpenXml.Presentation.ShapeTree,
        ByVal focus As EmuRect,
        ByRef bestReference As System.String,
        ByRef bestArea As System.Double)

        If tree Is Nothing Then Return
        For Each shp As DocumentFormat.OpenXml.Presentation.Shape In tree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
            Dim fillRef As System.String = ExtractSolidFillColor(shp.ShapeProperties)
            If System.String.IsNullOrWhiteSpace(fillRef) Then Continue For
            Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = shp.ShapeProperties?.Transform2D
            If xfrm?.Offset Is Nothing OrElse xfrm.Extents Is Nothing Then Continue For
            Dim candidate As EmuRect = TransformToRect(xfrm)
            If candidate.Width <= 0 OrElse candidate.Height <= 0 OrElse Not RectContainsCenter(candidate, focus) Then Continue For
            Dim area As System.Double = CDbl(candidate.Width) * CDbl(candidate.Height)
            If area < bestArea Then
                bestArea = area
                bestReference = fillRef
            End If
        Next
    End Sub

    Private Function ExtractLocalBackgroundColor(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal focusRect As System.Nullable(Of EmuRect)) As System.String

        Dim result As System.String = ExtractEffectiveBackgroundColor(sp)
        If Not focusRect.HasValue Then Return result
        Dim bestArea As System.Double = System.Double.MaxValue
        ConsiderBackgroundShapes(sp?.Slide?.CommonSlideData?.ShapeTree, focusRect.Value, result, bestArea)
        ConsiderBackgroundShapes(sp?.SlideLayoutPart?.SlideLayout?.CommonSlideData?.ShapeTree, focusRect.Value, result, bestArea)
        ConsiderBackgroundShapes(sp?.SlideLayoutPart?.SlideMasterPart?.SlideMaster?.CommonSlideData?.ShapeTree, focusRect.Value, result, bestArea)
        Return result
    End Function

    Private Function IsConsultingGeneratedBackground(ByVal value As System.String) As System.Boolean
        Dim c As System.String = NormalizeHexColor(value, "#FFFFFF")
        Return System.String.Equals(c, "#0B1F33", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(c, "#F6F8FA", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(c, "#EAF0F4", System.StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function IsDarkHexColor(ByVal value As System.String) As System.Boolean
        Dim c As System.String = NormalizeHexColor(value, "#FFFFFF")
        Dim r As System.Double = System.Convert.ToInt32(c.Substring(1, 2), 16) / 255.0R
        Dim g As System.Double = System.Convert.ToInt32(c.Substring(3, 2), 16) / 255.0R
        Dim b As System.Double = System.Convert.ToInt32(c.Substring(5, 2), 16) / 255.0R
        Dim luminance As System.Double = 0.2126R * r + 0.7152R * g + 0.0722R * b
        Return luminance < 0.48R
    End Function

    Private Function ResolveVisualRect(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement,
        Optional ByVal componentKind As System.String = "") As EmuRect

        Dim tf As System.Text.Json.JsonElement
        If el.TryGetProperty("transform", tf) AndAlso tf.ValueKind = System.Text.Json.JsonValueKind.Object Then
            Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetTransformFromJson(presPart, tf)
            Return TransformToRect(xfrm)
        End If

        Dim phEl As System.Text.Json.JsonElement
        If el.TryGetProperty("placeholder", phEl) Then
            Dim target As DocumentFormat.OpenXml.Presentation.Shape = FindShapeByPlaceholderElement(sp, phEl)

            ' Components must never be rendered into a title/subtitle/footer merely because
            ' an LLM supplied a stale shapeId copied from an example slide. Shape IDs are
            ' slide-local and can differ from the IDs on the selected layout.
            If target IsNot Nothing AndAlso Not IsSuitableComponentPlaceholder(sp, target) Then
                target = Nothing
            End If

            If target Is Nothing Then
                target = FindBestComponentPlaceholder(sp)
            End If

            If target IsNot Nothing Then
                Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, target)
                If xfrm IsNot Nothing Then
                    Return ExpandComponentRectForLayout(presPart, sp, TransformToRect(xfrm), componentKind)
                End If
            End If
        Else
            Dim target As DocumentFormat.OpenXml.Presentation.Shape = FindBestComponentPlaceholder(sp)
            If target IsNot Nothing Then
                Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, target)
                If xfrm IsNot Nothing Then
                    Return ExpandComponentRectForLayout(presPart, sp, TransformToRect(xfrm), componentKind)
                End If
            End If
        End If

        Dim w As Long = presPart.Presentation.SlideSize.Cx.Value
        Dim h As Long = presPart.Presentation.SlideSize.Cy.Value
        Return New EmuRect With {
            .X = CLng(w * 0.075R),
            .Y = CLng(h * 0.23R),
            .Width = CLng(w * 0.85R),
            .Height = CLng(h * 0.66R)
        }
    End Function

    Private Function ExpandComponentRectForLayout(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal requested As EmuRect,
        ByVal componentKind As System.String) As EmuRect

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return requested
        Dim kind As System.String = If(componentKind, System.String.Empty).Trim().ToLowerInvariant()
        If kind <> "comparison" AndAlso kind <> "compare" Then Return requested

        Dim rects As New System.Collections.Generic.List(Of EmuRect)()
        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In
            sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

            If Not IsSuitableComponentPlaceholder(sp, shape) Then Continue For
            Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shape)
            If xfrm Is Nothing Then Continue For
            Dim r As EmuRect = TransformToRect(xfrm)
            If r.Width > 0 AndAlso r.Height > 0 Then rects.Add(r)
        Next

        If rects.Count < 2 Then Return requested

        ' Two-content layouts are common in corporate templates. A comparison component
        ' needs the union of both content columns; otherwise the renderer creates two
        ' comparison cards inside the left column and wastes half the slide.
        Dim requestedCenterY As System.Double = requested.Y + requested.Height / 2.0R
        Dim band As System.Collections.Generic.List(Of EmuRect) =
            rects.
                Where(
                    Function(r)
                        Dim centerY As System.Double = r.Y + r.Height / 2.0R
                        Return System.Math.Abs(centerY - requestedCenterY) <= System.Math.Max(requested.Height, r.Height) * 0.22R
                    End Function).
                OrderBy(Function(r) r.X).
                ToList()

        If band.Count < 2 Then Return requested

        Dim left As Long = band.Min(Function(r) r.X)
        Dim top As Long = band.Min(Function(r) r.Y)
        Dim right As Long = band.Max(Function(r) r.X + r.Width)
        Dim bottom As Long = band.Max(Function(r) r.Y + r.Height)
        Dim unionRect As New EmuRect With {
            .X = left,
            .Y = top,
            .Width = right - left,
            .Height = bottom - top
        }

        If unionRect.Width > CLng(requested.Width * 1.25R) Then Return unionRect
        Return requested
    End Function

    Private Function GetTransformFromJson(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal tf As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Drawing.Transform2D

        Dim rawX As System.Double = 0.0R
        System.Double.TryParse(tf.GetProperty("x").GetRawText(), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, rawX)
        If rawX <= 1.0R Then Return ConvertRelativeToAbsoluteTransform(presPart, tf)
        Return New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset() With {.X = tf.GetProperty("x").GetInt64(), .Y = tf.GetProperty("y").GetInt64()},
            New DocumentFormat.OpenXml.Drawing.Extents() With {.Cx = tf.GetProperty("width").GetInt64(), .Cy = tf.GetProperty("height").GetInt64()})
    End Function

    Private Function TransformToRect(ByVal xfrm As DocumentFormat.OpenXml.Drawing.Transform2D) As EmuRect
        If xfrm Is Nothing Then Return New EmuRect()
        Return New EmuRect With {
            .X = If(xfrm.Offset?.X IsNot Nothing, xfrm.Offset.X.Value, 0L),
            .Y = If(xfrm.Offset?.Y IsNot Nothing, xfrm.Offset.Y.Value, 0L),
            .Width = If(xfrm.Extents?.Cx IsNot Nothing, xfrm.Extents.Cx.Value, 0L),
            .Height = If(xfrm.Extents?.Cy IsNot Nothing, xfrm.Extents.Cy.Value, 0L)
        }
    End Function

    Private Function RectTransform(ByVal rect As EmuRect) As DocumentFormat.OpenXml.Drawing.Transform2D
        Return New DocumentFormat.OpenXml.Drawing.Transform2D(
            New DocumentFormat.OpenXml.Drawing.Offset() With {.X = rect.X, .Y = rect.Y},
            New DocumentFormat.OpenXml.Drawing.Extents() With {.Cx = System.Math.Max(1L, rect.Width), .Cy = System.Math.Max(1L, rect.Height)})
    End Function

    Private Function NextShapeId(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As UInteger
        Return sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties)().
            Select(Function(nv) If(nv.Id IsNot Nothing, nv.Id.Value, 0UI)).DefaultIfEmpty(0UI).Max() + 1UI
    End Function

    Private Sub AddPrimitiveShape(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal rect As EmuRect,
        ByVal shapeType As DocumentFormat.OpenXml.Drawing.ShapeTypeValues,
        ByVal fillColor As System.String,
        Optional ByVal outlineColor As System.String = "",
        Optional ByVal outlinePoints As System.Double = 0.0R)

        Dim id As UInteger = NextShapeId(sp)
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = RectTransform(rect)}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {.Preset = shapeType})
        If System.String.IsNullOrWhiteSpace(fillColor) Then
            spPr.Append(New DocumentFormat.OpenXml.Drawing.NoFill())
        Else
            spPr.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = NormalizeHexColor(fillColor, "#FFFFFF").TrimStart("#"c)}))
        End If
        If System.String.IsNullOrWhiteSpace(outlineColor) OrElse outlinePoints <= 0.0R Then
            spPr.Append(New DocumentFormat.OpenXml.Drawing.Outline(New DocumentFormat.OpenXml.Drawing.NoFill()))
        Else
            Dim outline As New DocumentFormat.OpenXml.Drawing.Outline() With {.Width = CInt(outlinePoints * 12700.0R)}
            outline.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = NormalizeHexColor(outlineColor, "#CBD5E1").TrimStart("#"c)}))
            spPr.Append(outline)
        End If
        Dim nv As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = id, .Name = "Visual " & id.ToString(System.Globalization.CultureInfo.InvariantCulture)},
            New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())
        sp.Slide.CommonSlideData.ShapeTree.Append(New DocumentFormat.OpenXml.Presentation.Shape(nv, spPr))
    End Sub

    Private Function FitPrimitiveFontSize(
        ByVal rect As EmuRect,
        ByVal text As System.String,
        ByVal requestedSize As System.Double) As System.Double

        If requestedSize <= 0.0R OrElse System.String.IsNullOrWhiteSpace(text) Then Return requestedSize
        If rect.Width <= 0 OrElse rect.Height <= 0 Then Return requestedSize

        Dim widthPoints As System.Double = rect.Width / 12700.0R
        Dim heightPoints As System.Double = rect.Height / 12700.0R
        Dim minimumSize As System.Double = System.Math.Max(8.5R, requestedSize * 0.66R)
        Dim size As System.Double = requestedSize

        Do While size >= minimumSize
            Dim charsPerLine As System.Int32 =
                System.Math.Max(4, CInt(System.Math.Floor(widthPoints / System.Math.Max(1.0R, size * 0.53R))))
            Dim estimatedLines As System.Int32 = EstimateWrappedLineCount(text, charsPerLine)
            Dim maxLines As System.Int32 =
                System.Math.Max(1, CInt(System.Math.Floor(heightPoints / System.Math.Max(1.0R, size * 1.2R))))

            If estimatedLines <= maxLines Then Return size
            size -= 0.5R
        Loop

        Return minimumSize
    End Function

    Private Function EstimateWrappedLineCount(
        ByVal text As System.String,
        ByVal charsPerLine As System.Int32) As System.Int32

        If System.String.IsNullOrEmpty(text) Then Return 1
        charsPerLine = System.Math.Max(4, charsPerLine)
        Dim total As System.Int32 = 0
        Dim normalized As System.String = text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        For Each paragraph As System.String In normalized.Split(New System.String() {vbLf}, System.StringSplitOptions.None)
            If paragraph.Length = 0 Then
                total += 1
                Continue For
            End If

            Dim lineLength As System.Int32 = 0
            Dim lines As System.Int32 = 1
            For Each word As System.String In paragraph.Split(New System.Char() {" "c, Microsoft.VisualBasic.ChrW(9)}, System.StringSplitOptions.RemoveEmptyEntries)
                Dim wordLength As System.Int32 = word.Length
                If lineLength = 0 Then
                    lineLength = wordLength
                ElseIf lineLength + 1 + wordLength <= charsPerLine Then
                    lineLength += 1 + wordLength
                Else
                    lines += 1
                    lineLength = wordLength
                End If

                If wordLength > charsPerLine Then
                    Dim extra As System.Int32 = CInt(System.Math.Floor((wordLength - 1) / CDbl(charsPerLine)))
                    lines += extra
                    lineLength = wordLength Mod charsPerLine
                End If
            Next
            total += lines
        Next

        Return System.Math.Max(1, total)
    End Function

    Private Sub AddPrimitiveText(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal rect As EmuRect,
        ByVal text As System.String,
        ByVal font As System.String,
        ByVal fontSize As System.Double,
        ByVal color As System.String,
        ByVal bold As System.Boolean)

        AddPrimitiveText(
            sp,
            rect,
            text,
            font,
            fontSize,
            color,
            bold,
            DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Left)
    End Sub

    Private Sub AddPrimitiveText(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal rect As EmuRect,
        ByVal text As System.String,
        ByVal font As System.String,
        ByVal fontSize As System.Double,
        ByVal color As System.String,
        ByVal bold As System.Boolean,
        ByVal align As DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues)

        Dim id As UInteger = NextShapeId(sp)
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = RectTransform(rect)}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle})
        spPr.Append(New DocumentFormat.OpenXml.Drawing.NoFill())
        spPr.Append(New DocumentFormat.OpenXml.Drawing.Outline(New DocumentFormat.OpenXml.Drawing.NoFill()))

        Dim nvDr As New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties() With {.TextBox = True}
        Dim nv As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = id, .Name = "Text " & id.ToString(System.Globalization.CultureInfo.InvariantCulture)},
            nvDr,
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())

        Dim fittedFontSize As System.Double = FitPrimitiveFontSize(rect, text, fontSize)
        Dim rp As New DocumentFormat.OpenXml.Drawing.RunProperties() With {.FontSize = CInt(fittedFontSize * 100.0R), .Bold = bold}
        SetRunSolidColorOrdered(rp, color)
        If Not System.String.IsNullOrWhiteSpace(font) Then SetRunLatinFontOrdered(rp, font)
        Dim pPr As New DocumentFormat.OpenXml.Drawing.ParagraphProperties() With {.Alignment = align}
        pPr.Append(New DocumentFormat.OpenXml.Drawing.NoBullet())
        Dim paragraph As New DocumentFormat.OpenXml.Drawing.Paragraph(pPr, New DocumentFormat.OpenXml.Drawing.Run(rp, New DocumentFormat.OpenXml.Drawing.Text(If(text, System.String.Empty))))
        Dim tb As New DocumentFormat.OpenXml.Presentation.TextBody(New DocumentFormat.OpenXml.Drawing.BodyProperties(), New DocumentFormat.OpenXml.Drawing.ListStyle(), paragraph)
        sp.Slide.CommonSlideData.ShapeTree.Append(New DocumentFormat.OpenXml.Presentation.Shape(nv, spPr, tb))
    End Sub

    Private Sub AddPrimitiveLine(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal x1 As Long,
        ByVal y1 As Long,
        ByVal x2 As Long,
        ByVal y2 As Long,
        ByVal color As System.String,
        Optional ByVal widthPoints As System.Double = 1.5R)

        Dim rect As New EmuRect With {.X = System.Math.Min(x1, x2), .Y = System.Math.Min(y1, y2), .Width = System.Math.Abs(x2 - x1), .Height = System.Math.Abs(y2 - y1)}
        If rect.Width = 0 Then rect.Width = 1
        If rect.Height = 0 Then rect.Height = 1
        Dim id As UInteger = NextShapeId(sp)
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties() With {.Transform2D = RectTransform(rect)}
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Line})
        Dim outline As New DocumentFormat.OpenXml.Drawing.Outline() With {.Width = CInt(widthPoints * 12700.0R)}
        outline.Append(New DocumentFormat.OpenXml.Drawing.SolidFill(New DocumentFormat.OpenXml.Drawing.RgbColorModelHex() With {.Val = NormalizeHexColor(color, "#94A3B8").TrimStart("#"c)}))
        spPr.Append(outline)
        Dim nv As New DocumentFormat.OpenXml.Presentation.NonVisualShapeProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = id, .Name = "Line " & id.ToString(System.Globalization.CultureInfo.InvariantCulture)},
            New DocumentFormat.OpenXml.Presentation.NonVisualShapeDrawingProperties(),
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())
        sp.Slide.CommonSlideData.ShapeTree.Append(New DocumentFormat.OpenXml.Presentation.Shape(nv, spPr))
    End Sub

    Private Function JsonString(ByVal obj As System.Text.Json.JsonElement, ByVal propertyName As System.String, Optional ByVal fallback As System.String = "") As System.String
        Dim tmp As System.Text.Json.JsonElement
        If obj.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso obj.TryGetProperty(propertyName, tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then Return If(tmp.GetString(), fallback)
        Return fallback
    End Function

    Private Function JsonDouble(ByVal obj As System.Text.Json.JsonElement, ByVal propertyName As System.String, Optional ByVal fallback As System.Double = 0.0R) As System.Double
        Dim tmp As System.Text.Json.JsonElement
        If obj.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso obj.TryGetProperty(propertyName, tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then Return tmp.GetDouble()
        Return fallback
    End Function

    Private Function JsonInt(ByVal obj As System.Text.Json.JsonElement, ByVal propertyName As System.String, Optional ByVal fallback As System.Int32 = 0) As System.Int32
        Dim tmp As System.Text.Json.JsonElement
        If obj.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso obj.TryGetProperty(propertyName, tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then Return tmp.GetInt32()
        Return fallback
    End Function

    Private Function ReadNestedItems(ByVal obj As System.Text.Json.JsonElement, ByVal propertyName As System.String) As System.Collections.Generic.List(Of System.Text.Json.JsonElement)
        Dim answer As New System.Collections.Generic.List(Of System.Text.Json.JsonElement)()
        Dim tmp As System.Text.Json.JsonElement
        If obj.ValueKind <> System.Text.Json.JsonValueKind.Object OrElse Not obj.TryGetProperty(propertyName, tmp) OrElse tmp.ValueKind <> System.Text.Json.JsonValueKind.Array Then Return answer
        For Each item As System.Text.Json.JsonElement In tmp.EnumerateArray()
            answer.Add(item)
        Next
        Return answer
    End Function

    Private Function ReadStringArray(ByVal obj As System.Text.Json.JsonElement, ByVal propertyName As System.String) As System.Collections.Generic.List(Of System.String)
        Dim answer As New System.Collections.Generic.List(Of System.String)()
        Dim tmp As System.Text.Json.JsonElement
        If obj.ValueKind <> System.Text.Json.JsonValueKind.Object OrElse Not obj.TryGetProperty(propertyName, tmp) OrElse tmp.ValueKind <> System.Text.Json.JsonValueKind.Array Then Return answer
        For Each item As System.Text.Json.JsonElement In tmp.EnumerateArray()
            If item.ValueKind = System.Text.Json.JsonValueKind.String Then answer.Add(item.GetString())
        Next
        Return answer
    End Function

    Private Sub RemoveComponentPlaceholders(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement,
        ByVal componentRect As EmuRect,
        ByVal componentKind As System.String)

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return

        Dim candidates As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
            sp.Slide.CommonSlideData.ShapeTree.
                Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)().
                Where(Function(shape) IsSuitableComponentPlaceholder(sp, shape)).
                ToList()

        For Each target As DocumentFormat.OpenXml.Presentation.Shape In candidates
            Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, target)
            If xfrm Is Nothing Then Continue For
            Dim r As EmuRect = TransformToRect(xfrm)
            If RectOverlapRatio(r, componentRect) >= 0.5R Then target.Remove()
        Next
    End Sub

    Private Function RectOverlapRatio(ByVal candidate As EmuRect, ByVal focus As EmuRect) As System.Double
        If candidate.Width <= 0 OrElse candidate.Height <= 0 OrElse focus.Width <= 0 OrElse focus.Height <= 0 Then Return 0.0R
        Dim left As Long = System.Math.Max(candidate.X, focus.X)
        Dim top As Long = System.Math.Max(candidate.Y, focus.Y)
        Dim right As Long = System.Math.Min(candidate.X + candidate.Width, focus.X + focus.Width)
        Dim bottom As Long = System.Math.Min(candidate.Y + candidate.Height, focus.Y + focus.Height)
        If right <= left OrElse bottom <= top Then Return 0.0R
        Dim overlap As System.Double = CDbl(right - left) * CDbl(bottom - top)
        Dim candidateArea As System.Double = CDbl(candidate.Width) * CDbl(candidate.Height)
        Return overlap / candidateArea
    End Function

    Private Function IsSuitableComponentPlaceholder(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal shape As DocumentFormat.OpenXml.Presentation.Shape) As System.Boolean

        If shape Is Nothing Then Return False
        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape =
            shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
        If ph Is Nothing Then Return False

        Dim semanticRole As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shape)

        ' Corporate templates frequently encode eyebrow/kicker strips as p:ph type="body".
        ' They are semantically body-like to PowerPoint but are far too small to host a chart,
        ' cards or another infographic. Only the real content body is a component canvas.
        If Not System.String.Equals(semanticRole, "body", System.StringComparison.OrdinalIgnoreCase) Then Return False

        Return IsBodyLikePlaceholder(ph)
    End Function

    Private Function FindBestComponentPlaceholder(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart) As DocumentFormat.OpenXml.Presentation.Shape

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return Nothing

        Dim candidates As New System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape)()
        For Each shape As DocumentFormat.OpenXml.Presentation.Shape In
            sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

            If IsSuitableComponentPlaceholder(sp, shape) Then candidates.Add(shape)
        Next

        If candidates.Count = 0 Then Return Nothing

        ' Prefer the largest body/object placeholder because that is the safest content
        ' canvas for charts and infographics on an arbitrary corporate layout.
        Return candidates.
            OrderByDescending(
                Function(shape)
                    Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shape)
                    If xfrm?.Extents Is Nothing OrElse xfrm.Extents.Cx Is Nothing OrElse xfrm.Extents.Cy Is Nothing Then Return 0.0R
                    Return CDbl(xfrm.Extents.Cx.Value) * CDbl(xfrm.Extents.Cy.Value)
                End Function).
            FirstOrDefault()
    End Function

    Private Sub ApplyComponentStyleOverrides(ByVal ctx As RenderStyleContext, ByVal el As System.Text.Json.JsonElement)
        If ctx Is Nothing Then Return
        Dim styleEl As System.Text.Json.JsonElement
        If Not el.TryGetProperty("style", styleEl) OrElse styleEl.ValueKind <> System.Text.Json.JsonValueKind.Object Then Return
        Dim value As System.String
        value = JsonString(styleEl, "accentColor", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.Accent1 = NormalizeHexColor(value, ctx.Accent1)
        value = JsonString(styleEl, "accent2", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.Accent2 = NormalizeHexColor(value, ctx.Accent2)
        value = JsonString(styleEl, "accent3", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.Accent3 = NormalizeHexColor(value, ctx.Accent3)
        value = JsonString(styleEl, "textColor", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.Dark = NormalizeHexColor(value, ctx.Dark)
        value = JsonString(styleEl, "surfaceColor", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.Surface = NormalizeHexColor(value, ctx.Surface)
        value = JsonString(styleEl, "headingFont", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.HeadingFont = value
        value = JsonString(styleEl, "bodyFont", "")
        If Not System.String.IsNullOrWhiteSpace(value) Then ctx.BodyFont = value
        ctx.MutedSurface = MixHexColors(ctx.Accent1, ctx.Light, 0.9R)
        ctx.MutedText = MixHexColors(ctx.Dark, ctx.Light, 0.35R)
    End Sub

    Private Sub AddComponent(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement)

        Dim kind As System.String = JsonString(el, "kind", "cards").Trim().ToLowerInvariant()
        Dim rect As EmuRect = ResolveVisualRect(presPart, sp, el, kind)
        Dim ctx As RenderStyleContext = GetRenderStyleContext(sp, rect)
        ApplyComponentStyleOverrides(ctx, el)
        RemoveComponentPlaceholders(sp, el, rect, kind)

        Select Case kind
            Case "metric_cards", "metrics", "kpi_cards"
                RenderMetricCards(sp, el, rect, ctx)
            Case "process", "steps"
                RenderProcess(sp, el, rect, ctx)
            Case "timeline"
                RenderTimeline(sp, el, rect, ctx)
            Case "comparison", "compare"
                RenderComparison(sp, el, rect, ctx)
            Case "bar_chart", "bars", "horizontal_bar"
                RenderBarChart(sp, el, rect, ctx)
            Case "column_chart", "columns"
                RenderColumnChart(sp, el, rect, ctx)
            Case "stacked_bar", "stacked_bars"
                RenderStackedBarChart(sp, el, rect, ctx)
            Case "line_chart", "lines"
                RenderLineChart(sp, el, rect, ctx)
            Case "waterfall", "bridge"
                RenderWaterfall(sp, el, rect, ctx)
            Case "dot_plot", "benchmark_plot"
                RenderDotPlot(sp, el, rect, ctx)
            Case "driver_tree", "issue_tree"
                RenderDriverTree(sp, el, rect, ctx)
            Case "agenda"
                RenderAgenda(sp, el, rect, ctx)
            Case "big_number", "big_metric"
                RenderBigNumber(sp, el, rect, ctx)
            Case "quote"
                RenderQuote(sp, el, rect, ctx)
            Case "callout"
                RenderCallout(sp, el, rect, ctx)
            Case "matrix", "pillars", "cards"
                RenderGenericCards(sp, el, rect, ctx)
            Case Else
                RenderGenericCards(sp, el, rect, ctx)
        End Select
        sp.Slide.Save()
    End Sub

    Private Function GetComponentItems(ByVal el As System.Text.Json.JsonElement) As System.Collections.Generic.List(Of System.Text.Json.JsonElement)
        Dim items As New System.Collections.Generic.List(Of System.Text.Json.JsonElement)()
        Dim tmp As System.Text.Json.JsonElement
        If el.TryGetProperty("items", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Array Then
            For Each item As System.Text.Json.JsonElement In tmp.EnumerateArray()
                items.Add(item)
            Next
        End If
        Return items
    End Function

    Private Function GetComponentDetail(ByVal item As System.Text.Json.JsonElement) As System.String
        Dim value As System.String = JsonString(item, "detail", System.String.Empty)
        If Not System.String.IsNullOrWhiteSpace(value) Then Return value
        value = JsonString(item, "text", System.String.Empty)
        If Not System.String.IsNullOrWhiteSpace(value) Then Return value
        value = JsonString(item, "description", System.String.Empty)
        If Not System.String.IsNullOrWhiteSpace(value) Then Return value
        value = JsonString(item, "subtitle", System.String.Empty)
        Return value
    End Function

    Private Function GetComponentPoints(ByVal item As System.Text.Json.JsonElement) As System.Collections.Generic.List(Of System.String)
        Dim result As System.Collections.Generic.List(Of System.String) = ReadStringArray(item, "points")
        If result.Count > 0 Then Return result
        result = ReadStringArray(item, "bullets")
        If result.Count > 0 Then Return result
        result = ReadStringArray(item, "items")
        Return result
    End Function

    Private Sub RenderMetricCards(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(4).ToList()
        If items.Count = 0 Then Return
        Dim gap As Long = CLng(rect.Width * 0.018R)
        Dim cardW As Long = CLng((rect.Width - gap * (items.Count - 1)) / CDbl(items.Count))
        Dim accents() As System.String = {ctx.Accent1, ctx.Accent2, ctx.Accent3, ctx.Accent1}
        For i As System.Int32 = 0 To items.Count - 1
            Dim x As Long = rect.X + i * (cardW + gap)
            Dim card As New EmuRect With {.X = x, .Y = rect.Y, .Width = cardW, .Height = rect.Height}
            AddPrimitiveShape(sp, card, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle, ctx.Surface, MixHexColors(ctx.Dark, ctx.Light, 0.82R), 0.7R)
            AddPrimitiveShape(sp, New EmuRect With {.X = x, .Y = rect.Y, .Width = cardW, .Height = System.Math.Max(70000L, CLng(rect.Height * 0.035R))}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, accents(i))
            Dim pad As Long = CLng(cardW * 0.08R)
            Dim value As System.String = JsonString(items(i), "value", JsonString(items(i), "number", ""))
            Dim label As System.String = JsonString(items(i), "label", JsonString(items(i), "title", ""))
            Dim detail As System.String = GetComponentDetail(items(i))
            AddPrimitiveText(sp, New EmuRect With {.X = x + pad, .Y = rect.Y + CLng(rect.Height * 0.17R), .Width = cardW - pad * 2, .Height = CLng(rect.Height * 0.25R)}, value, ctx.HeadingFont, System.Math.Max(22.0R, ctx.HeadingSize * 1.55R), ctx.Dark, True)
            AddPrimitiveText(sp, New EmuRect With {.X = x + pad, .Y = rect.Y + CLng(rect.Height * 0.47R), .Width = cardW - pad * 2, .Height = CLng(rect.Height * 0.15R)}, label, ctx.BodyFont, ctx.BodySize, ctx.Dark, True)
            If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = x + pad, .Y = rect.Y + CLng(rect.Height * 0.66R), .Width = cardW - pad * 2, .Height = CLng(rect.Height * 0.2R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
        Next
    End Sub

    Private Sub RenderProcess(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(6).ToList()
        If items.Count = 0 Then Return
        Dim centerY As Long = rect.Y + CLng(rect.Height * 0.3R)
        Dim left As Long = rect.X + CLng(rect.Width * 0.07R)
        Dim right As Long = rect.X + rect.Width - CLng(rect.Width * 0.07R)
        If items.Count > 1 Then AddPrimitiveLine(sp, left, centerY, right, centerY, MixHexColors(ctx.Dark, ctx.Light, 0.72R), 1.4R)
        Dim stepW As Long = CLng(rect.Width / CDbl(items.Count))
        For i As System.Int32 = 0 To items.Count - 1
            Dim cx As Long = rect.X + CLng(stepW * (i + 0.5R))
            Dim dot As Long = System.Math.Min(CLng(rect.Height * 0.16R), CLng(stepW * 0.32R))
            AddPrimitiveShape(sp, New EmuRect With {.X = cx - dot \ 2, .Y = centerY - dot \ 2, .Width = dot, .Height = dot}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, If(i Mod 2 = 0, ctx.Accent1, ctx.Accent2))
            AddPrimitiveText(sp, New EmuRect With {.X = cx - dot \ 2, .Y = centerY - CLng(dot * 0.32R), .Width = dot, .Height = CLng(dot * 0.55R)}, (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), ctx.HeadingFont, System.Math.Max(11.0R, ctx.SmallSize), "#FFFFFF", True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * stepW + CLng(stepW * 0.07R), .Y = rect.Y + CLng(rect.Height * 0.49R), .Width = CLng(stepW * 0.86R), .Height = CLng(rect.Height * 0.16R)}, JsonString(items(i), "title", JsonString(items(i), "label", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            Dim detail As System.String = GetComponentDetail(items(i))
            If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * stepW + CLng(stepW * 0.08R), .Y = rect.Y + CLng(rect.Height * 0.67R), .Width = CLng(stepW * 0.84R), .Height = CLng(rect.Height * 0.21R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
        Next
    End Sub

    Private Sub RenderTimeline(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(6).ToList()
        If items.Count = 0 Then Return
        Dim lineY As Long = rect.Y + CLng(rect.Height * 0.48R)
        Dim marginX As Long = CLng(rect.Width * 0.07R)
        AddPrimitiveLine(sp, rect.X + marginX, lineY, rect.X + rect.Width - marginX, lineY, MixHexColors(ctx.Dark, ctx.Light, 0.7R), 1.5R)
        Dim slotW As Long = CLng((rect.Width - 2 * marginX) / CDbl(System.Math.Max(1, items.Count)))
        For i As System.Int32 = 0 To items.Count - 1
            Dim cx As Long = rect.X + marginX + CLng(slotW * (i + 0.5R))
            Dim dot As Long = CLng(rect.Height * 0.075R)
            AddPrimitiveShape(sp, New EmuRect With {.X = cx - dot \ 2, .Y = lineY - dot \ 2, .Width = dot, .Height = dot}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, If(i Mod 2 = 0, ctx.Accent1, ctx.Accent2))
            AddPrimitiveText(sp, New EmuRect With {.X = cx - CLng(slotW * 0.45R), .Y = rect.Y + CLng(rect.Height * 0.1R), .Width = CLng(slotW * 0.9R), .Height = CLng(rect.Height * 0.14R)}, JsonString(items(i), "date", JsonString(items(i), "label", "")), ctx.BodyFont, ctx.SmallSize, If(i Mod 2 = 0, ctx.Accent1, ctx.Accent2), True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            AddPrimitiveText(sp, New EmuRect With {.X = cx - CLng(slotW * 0.45R), .Y = rect.Y + CLng(rect.Height * 0.61R), .Width = CLng(slotW * 0.9R), .Height = CLng(rect.Height * 0.14R)}, JsonString(items(i), "title", ""), ctx.BodyFont, ctx.BodySize, ctx.Dark, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            Dim detail As System.String = GetComponentDetail(items(i))
            If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = cx - CLng(slotW * 0.45R), .Y = rect.Y + CLng(rect.Height * 0.76R), .Width = CLng(slotW * 0.9R), .Height = CLng(rect.Height * 0.17R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
        Next
    End Sub

    Private Sub RenderComparison(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement,
        ByVal rect As EmuRect,
        ByVal ctx As RenderStyleContext)

        Dim items = GetComponentItems(el).Take(2).ToList()
        If items.Count = 0 Then Return

        Dim gap As Long = CLng(rect.Width * 0.028R)
        Dim colW As Long = CLng((rect.Width - gap) / 2.0R)
        Dim headerHeight As Long = CLng(rect.Height * 0.19R)
        Dim bodyTop As Long = rect.Y + CLng(rect.Height * 0.245R)
        Dim bodyBottom As Long = rect.Y + CLng(rect.Height * 0.94R)
        Dim bodyHeight As Long = System.Math.Max(1L, bodyBottom - bodyTop)

        For i As System.Int32 = 0 To items.Count - 1
            Dim x As Long = rect.X + i * (colW + gap)
            Dim accent As System.String = If(i = 0, ctx.Accent1, ctx.Accent2)

            AddPrimitiveShape(
                sp,
                New EmuRect With {.X = x, .Y = rect.Y, .Width = colW, .Height = rect.Height},
                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle,
                ctx.Surface,
                MixHexColors(ctx.Dark, ctx.Light, 0.82R),
                0.7R)

            AddPrimitiveShape(
                sp,
                New EmuRect With {.X = x, .Y = rect.Y, .Width = colW, .Height = headerHeight},
                DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle,
                MixHexColors(accent, ctx.Light, 0.84R))

            AddPrimitiveText(
                sp,
                New EmuRect With {
                    .X = x + CLng(colW * 0.065R),
                    .Y = rect.Y + CLng(headerHeight * 0.18R),
                    .Width = CLng(colW * 0.87R),
                    .Height = CLng(headerHeight * 0.64R)},
                JsonString(items(i), "title", ""),
                ctx.HeadingFont,
                System.Math.Min(ctx.HeadingSize, 16.0R),
                accent,
                True)

            Dim points = GetComponentPoints(items(i)).Take(5).ToList()
            If points.Count = 0 Then Continue For

            Dim rowH As Long = CLng(bodyHeight / CDbl(points.Count))
            For j As System.Int32 = 0 To points.Count - 1
                Dim y As Long = bodyTop + j * rowH
                Dim dot As Long = System.Math.Max(55000L, System.Math.Min(CLng(rect.Height * 0.023R), CLng(rowH * 0.16R)))
                Dim dotY As Long = y + CLng(rowH * 0.14R)
                AddPrimitiveShape(
                    sp,
                    New EmuRect With {
                        .X = x + CLng(colW * 0.065R),
                        .Y = dotY,
                        .Width = dot,
                        .Height = dot},
                    DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse,
                    accent)

                AddPrimitiveText(
                    sp,
                    New EmuRect With {
                        .X = x + CLng(colW * 0.13R),
                        .Y = y + CLng(rowH * 0.04R),
                        .Width = CLng(colW * 0.79R),
                        .Height = CLng(rowH * 0.84R)},
                    points(j),
                    ctx.BodyFont,
                    System.Math.Min(ctx.SmallSize, 11.5R),
                    ctx.Dark,
                    False)
            Next
        Next
    End Sub

    Private Sub RenderBarChart(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(8).ToList()
        If items.Count = 0 Then Return

        Dim values As System.Collections.Generic.List(Of System.Double) = items.Select(Function(i) JsonDouble(i, "value", 0.0R)).ToList()
        Dim minValue As System.Double = System.Math.Min(0.0R, values.Min())
        Dim maxValue As System.Double = System.Math.Max(0.0R, values.Max())
        If System.Math.Abs(maxValue - minValue) < 0.000001R Then maxValue = minValue + 1.0R

        Dim highlightIndex As System.Int32 = JsonInt(el, "highlightIndex", -1)
        If highlightIndex < 0 OrElse highlightIndex >= items.Count Then
            highlightIndex = values.IndexOf(values.Max())
        End If

        Dim rowH As Long = CLng(rect.Height / CDbl(items.Count))
        Dim labelW As Long = CLng(rect.Width * 0.25R)
        Dim valueW As Long = CLng(rect.Width * 0.13R)
        Dim chartX As Long = rect.X + labelW + CLng(rect.Width * 0.02R)
        Dim chartW As Long = rect.Width - labelW - valueW - CLng(rect.Width * 0.05R)
        Dim zeroX As Long = chartX + CLng(chartW * ((0.0R - minValue) / (maxValue - minValue)))

        AddPrimitiveLine(sp, zeroX, rect.Y, zeroX, rect.Y + rect.Height, MixHexColors(ctx.Dark, ctx.Background, 0.72R), 0.7R)

        For i As System.Int32 = 0 To items.Count - 1
            Dim y As Long = rect.Y + i * rowH
            Dim value As System.Double = values(i)
            AddPrimitiveText(sp,
                New EmuRect With {.X = rect.X, .Y = y + CLng(rowH * 0.13R), .Width = labelW - CLng(rect.Width * 0.01R), .Height = CLng(rowH * 0.58R)},
                JsonString(items(i), "label", JsonString(items(i), "title", "")),
                ctx.BodyFont, ctx.BodySize, ctx.Dark, i = highlightIndex)

            Dim valueX As Long = chartX + CLng(chartW * ((value - minValue) / (maxValue - minValue)))
            Dim barLeft As Long = System.Math.Min(zeroX, valueX)
            Dim barW As Long = System.Math.Max(1L, System.Math.Abs(valueX - zeroX))
            Dim barH As Long = System.Math.Max(45000L, CLng(rowH * 0.34R))
            Dim barY As Long = y + CLng(rowH * 0.28R)
            Dim barColor As System.String = If(i = highlightIndex, ctx.Accent1, MixHexColors(ctx.Accent1, ctx.Background, 0.58R))
            AddPrimitiveShape(sp, New EmuRect With {.X = barLeft, .Y = barY, .Width = barW, .Height = barH}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, barColor)

            AddPrimitiveText(sp,
                New EmuRect With {.X = chartX + chartW + CLng(rect.Width * 0.015R), .Y = y + CLng(rowH * 0.13R), .Width = valueW, .Height = CLng(rowH * 0.58R)},
                JsonString(items(i), "displayValue", value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)),
                ctx.BodyFont, ctx.BodySize, ctx.Dark, i = highlightIndex, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right)
        Next
    End Sub

    Private Sub RenderColumnChart(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(9).ToList()
        If items.Count = 0 Then Return

        Dim values As System.Collections.Generic.List(Of System.Double) = items.Select(Function(i) JsonDouble(i, "value", 0.0R)).ToList()
        Dim minValue As System.Double = System.Math.Min(0.0R, values.Min())
        Dim maxValue As System.Double = System.Math.Max(0.0R, values.Max())
        If System.Math.Abs(maxValue - minValue) < 0.000001R Then maxValue = minValue + 1.0R

        Dim highlightIndex As System.Int32 = JsonInt(el, "highlightIndex", values.IndexOf(values.Max()))
        Dim labelZone As Long = CLng(rect.Height * 0.18R)
        Dim valueZone As Long = CLng(rect.Height * 0.12R)
        Dim plotY As Long = rect.Y + valueZone
        Dim plotH As Long = rect.Height - labelZone - valueZone
        Dim zeroY As Long = plotY + CLng(plotH * (maxValue / (maxValue - minValue)))
        AddPrimitiveLine(sp, rect.X, zeroY, rect.X + rect.Width, zeroY, MixHexColors(ctx.Dark, ctx.Background, 0.72R), 0.7R)

        Dim slotW As Long = CLng(rect.Width / CDbl(items.Count))
        Dim barW As Long = CLng(slotW * 0.48R)
        For i As System.Int32 = 0 To items.Count - 1
            Dim value As System.Double = values(i)
            Dim x As Long = rect.X + i * slotW + (slotW - barW) \ 2
            Dim valueY As Long = plotY + CLng(plotH * ((maxValue - value) / (maxValue - minValue)))
            Dim top As Long = System.Math.Min(zeroY, valueY)
            Dim h As Long = System.Math.Max(1L, System.Math.Abs(zeroY - valueY))
            Dim color As System.String = If(i = highlightIndex, ctx.Accent1, MixHexColors(ctx.Accent1, ctx.Background, 0.6R))
            AddPrimitiveShape(sp, New EmuRect With {.X = x, .Y = top, .Width = barW, .Height = h}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, color)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * slotW, .Y = rect.Y, .Width = slotW, .Height = valueZone}, JsonString(items(i), "displayValue", value.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)), ctx.BodyFont, ctx.SmallSize, ctx.Dark, i = highlightIndex, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * slotW, .Y = rect.Y + rect.Height - labelZone + CLng(labelZone * 0.12R), .Width = slotW, .Height = CLng(labelZone * 0.68R)}, JsonString(items(i), "label", JsonString(items(i), "title", "")), ctx.BodyFont, ctx.SmallSize, ctx.Dark, False, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
        Next
    End Sub

    Private Sub RenderStackedBarChart(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(7).ToList()
        If items.Count = 0 Then Return

        Dim totals As New System.Collections.Generic.List(Of System.Double)()
        For Each item As System.Text.Json.JsonElement In items
            totals.Add(ReadNestedItems(item, "segments").Sum(Function(s) System.Math.Max(0.0R, JsonDouble(s, "value", 0.0R))))
        Next
        Dim maxTotal As System.Double = totals.DefaultIfEmpty(1.0R).Max()
        If maxTotal <= 0.0R Then maxTotal = 1.0R

        Dim rowH As Long = CLng(rect.Height / CDbl(items.Count))
        Dim labelW As Long = CLng(rect.Width * 0.24R)
        Dim valueW As Long = CLng(rect.Width * 0.12R)
        Dim barX As Long = rect.X + labelW + CLng(rect.Width * 0.02R)
        Dim barAreaW As Long = rect.Width - labelW - valueW - CLng(rect.Width * 0.05R)
        Dim palette() As System.String = {
            ctx.Accent1,
            ctx.Accent2,
            ctx.Accent3,
            MixHexColors(ctx.Accent1, ctx.Background, 0.48R),
            MixHexColors(ctx.Accent2, ctx.Background, 0.48R)
        }

        For i As System.Int32 = 0 To items.Count - 1
            Dim y As Long = rect.Y + i * rowH
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X, .Y = y + CLng(rowH * 0.12R), .Width = labelW, .Height = CLng(rowH * 0.58R)}, JsonString(items(i), "label", JsonString(items(i), "title", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, False)
            Dim segments = ReadNestedItems(items(i), "segments")
            Dim cursor As Long = barX
            For j As System.Int32 = 0 To segments.Count - 1
                Dim v As System.Double = System.Math.Max(0.0R, JsonDouble(segments(j), "value", 0.0R))
                Dim w As Long = CLng(barAreaW * (v / maxTotal))
                If w <= 0 Then Continue For
                AddPrimitiveShape(sp, New EmuRect With {.X = cursor, .Y = y + CLng(rowH * 0.29R), .Width = w, .Height = CLng(rowH * 0.32R)}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, palette(j Mod palette.Length))
                cursor += w
            Next
            AddPrimitiveText(sp, New EmuRect With {.X = barX + barAreaW + CLng(rect.Width * 0.015R), .Y = y + CLng(rowH * 0.12R), .Width = valueW, .Height = CLng(rowH * 0.58R)}, JsonString(items(i), "displayValue", totals(i).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)), ctx.BodyFont, ctx.BodySize, ctx.Dark, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right)
        Next
    End Sub

    Private Sub RenderLineChart(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(10).ToList()
        If items.Count < 2 Then Return

        Dim values1 As System.Collections.Generic.List(Of System.Double) = items.Select(Function(i) JsonDouble(i, "value", 0.0R)).ToList()
        Dim hasSecond As System.Boolean = items.Any(Function(i)
                                                        Dim tmp As System.Text.Json.JsonElement
                                                        Return i.ValueKind = System.Text.Json.JsonValueKind.Object AndAlso i.TryGetProperty("value2", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number
                                                    End Function)
        Dim values2 As System.Collections.Generic.List(Of System.Double) = If(hasSecond, items.Select(Function(i) JsonDouble(i, "value2", 0.0R)).ToList(), New System.Collections.Generic.List(Of System.Double)())

        Dim allValues As New System.Collections.Generic.List(Of System.Double)(values1)
        If hasSecond Then allValues.AddRange(values2)
        Dim minValue As System.Double = allValues.Min()
        Dim maxValue As System.Double = allValues.Max()
        If System.Math.Abs(maxValue - minValue) < 0.000001R Then
            minValue -= 0.5R
            maxValue += 0.5R
        Else
            Dim pad As System.Double = (maxValue - minValue) * 0.08R
            minValue -= pad
            maxValue += pad
        End If

        Dim labelZone As Long = CLng(rect.Height * 0.18R)
        Dim plot As New EmuRect With {.X = rect.X + CLng(rect.Width * 0.02R), .Y = rect.Y + CLng(rect.Height * 0.06R), .Width = CLng(rect.Width * 0.92R), .Height = rect.Height - labelZone - CLng(rect.Height * 0.08R)}
        AddPrimitiveLine(sp, plot.X, plot.Y + plot.Height, plot.X + plot.Width, plot.Y + plot.Height, MixHexColors(ctx.Dark, ctx.Background, 0.75R), 0.7R)

        Dim points1 As New System.Collections.Generic.List(Of System.Tuple(Of Long, Long))()
        Dim points2 As New System.Collections.Generic.List(Of System.Tuple(Of Long, Long))()
        For i As System.Int32 = 0 To items.Count - 1
            Dim x As Long = plot.X + CLng(plot.Width * (i / CDbl(items.Count - 1)))
            Dim y1 As Long = plot.Y + CLng(plot.Height * ((maxValue - values1(i)) / (maxValue - minValue)))
            points1.Add(System.Tuple.Create(x, y1))
            If hasSecond Then
                Dim y2 As Long = plot.Y + CLng(plot.Height * ((maxValue - values2(i)) / (maxValue - minValue)))
                points2.Add(System.Tuple.Create(x, y2))
            End If
            AddPrimitiveText(sp, New EmuRect With {.X = x - CLng(plot.Width / CDbl(items.Count) * 0.5R), .Y = rect.Y + rect.Height - labelZone + CLng(labelZone * 0.12R), .Width = CLng(plot.Width / CDbl(items.Count)), .Height = CLng(labelZone * 0.58R)}, JsonString(items(i), "label", JsonString(items(i), "date", "")), ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
        Next

        DrawLineSeries(sp, points1, ctx.Accent1, ctx)
        If hasSecond Then DrawLineSeries(sp, points2, ctx.Accent2, ctx)

        Dim last1 = points1(points1.Count - 1)
        AddPrimitiveText(sp, New EmuRect With {.X = last1.Item1 - CLng(rect.Width * 0.1R), .Y = last1.Item2 - CLng(rect.Height * 0.08R), .Width = CLng(rect.Width * 0.1R), .Height = CLng(rect.Height * 0.07R)}, JsonString(items(items.Count - 1), "displayValue", values1(values1.Count - 1).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)), ctx.BodyFont, ctx.SmallSize, ctx.Accent1, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right)
        If hasSecond Then
            Dim last2 = points2(points2.Count - 1)
            AddPrimitiveText(sp, New EmuRect With {.X = last2.Item1 - CLng(rect.Width * 0.1R), .Y = last2.Item2 + CLng(rect.Height * 0.01R), .Width = CLng(rect.Width * 0.1R), .Height = CLng(rect.Height * 0.07R)}, JsonString(items(items.Count - 1), "displayValue2", values2(values2.Count - 1).ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)), ctx.BodyFont, ctx.SmallSize, ctx.Accent2, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Right)
        End If
    End Sub

    Private Sub DrawLineSeries(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal points As System.Collections.Generic.List(Of System.Tuple(Of Long, Long)),
        ByVal color As System.String,
        ByVal ctx As RenderStyleContext)

        If points Is Nothing OrElse points.Count < 2 Then Return
        For i As System.Int32 = 1 To points.Count - 1
            AddPrimitiveLine(sp, points(i - 1).Item1, points(i - 1).Item2, points(i).Item1, points(i).Item2, color, 1.8R)
        Next
        Dim dot As Long = 70000L
        For Each p As System.Tuple(Of Long, Long) In points
            AddPrimitiveShape(sp, New EmuRect With {.X = p.Item1 - dot \ 2, .Y = p.Item2 - dot \ 2, .Width = dot, .Height = dot}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, color)
        Next
    End Sub

    Private Sub RenderWaterfall(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(9).ToList()
        If items.Count = 0 Then Return

        Dim starts As New System.Collections.Generic.List(Of System.Double)()
        Dim ends As New System.Collections.Generic.List(Of System.Double)()
        Dim cumulative As System.Double = 0.0R
        For Each item As System.Text.Json.JsonElement In items
            Dim kind As System.String = JsonString(item, "kind", "delta").Trim().ToLowerInvariant()
            Dim value As System.Double = JsonDouble(item, "value", 0.0R)
            If kind = "total" OrElse kind = "subtotal" Then
                starts.Add(0.0R)
                ends.Add(value)
                cumulative = value
            Else
                starts.Add(cumulative)
                cumulative += value
                ends.Add(cumulative)
            End If
        Next

        Dim minValue As System.Double = System.Math.Min(0.0R, starts.Concat(ends).Min())
        Dim maxValue As System.Double = System.Math.Max(0.0R, starts.Concat(ends).Max())
        If System.Math.Abs(maxValue - minValue) < 0.000001R Then maxValue = minValue + 1.0R

        Dim labelZone As Long = CLng(rect.Height * 0.2R)
        Dim valueZone As Long = CLng(rect.Height * 0.12R)
        Dim plotY As Long = rect.Y + valueZone
        Dim plotH As Long = rect.Height - labelZone - valueZone
        Dim slotW As Long = CLng(rect.Width / CDbl(items.Count))
        Dim barW As Long = CLng(slotW * 0.52R)

        For i As System.Int32 = 0 To items.Count - 1
            Dim x As Long = rect.X + i * slotW + (slotW - barW) \ 2
            Dim yStart As Long = plotY + CLng(plotH * ((maxValue - starts(i)) / (maxValue - minValue)))
            Dim yEnd As Long = plotY + CLng(plotH * ((maxValue - ends(i)) / (maxValue - minValue)))
            Dim top As Long = System.Math.Min(yStart, yEnd)
            Dim h As Long = System.Math.Max(1L, System.Math.Abs(yStart - yEnd))
            Dim kind As System.String = JsonString(items(i), "kind", "delta").Trim().ToLowerInvariant()
            Dim delta As System.Double = ends(i) - starts(i)
            Dim color As System.String
            If kind = "total" OrElse kind = "subtotal" Then
                color = ctx.Accent1
            ElseIf delta >= 0.0R Then
                color = ctx.Accent2
            Else
                color = ctx.Accent3
            End If
            AddPrimitiveShape(sp, New EmuRect With {.X = x, .Y = top, .Width = barW, .Height = h}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, color)

            If i < items.Count - 1 Then
                Dim connectorY As Long = yEnd
                AddPrimitiveLine(sp, x + barW, connectorY, rect.X + (i + 1) * slotW + (slotW - barW) \ 2, connectorY, MixHexColors(ctx.Dark, ctx.Background, 0.7R), 0.7R)
            End If

            Dim rawValue As System.Double = JsonDouble(items(i), "value", delta)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * slotW, .Y = rect.Y, .Width = slotW, .Height = valueZone}, JsonString(items(i), "displayValue", rawValue.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture)), ctx.BodyFont, ctx.SmallSize, ctx.Dark, True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + i * slotW, .Y = rect.Y + rect.Height - labelZone + CLng(labelZone * 0.1R), .Width = slotW, .Height = CLng(labelZone * 0.66R)}, JsonString(items(i), "label", JsonString(items(i), "title", "")), ctx.BodyFont, ctx.SmallSize, ctx.Dark, False, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
        Next
    End Sub

    Private Sub RenderDotPlot(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(8).ToList()
        If items.Count = 0 Then Return

        Dim allValues As New System.Collections.Generic.List(Of System.Double)()
        For Each item As System.Text.Json.JsonElement In items
            allValues.Add(JsonDouble(item, "value", 0.0R))
            allValues.Add(JsonDouble(item, "benchmark", JsonDouble(item, "value2", 0.0R)))
        Next
        Dim minValue As System.Double = allValues.Min()
        Dim maxValue As System.Double = allValues.Max()
        If System.Math.Abs(maxValue - minValue) < 0.000001R Then maxValue = minValue + 1.0R
        Dim pad As System.Double = (maxValue - minValue) * 0.08R
        minValue -= pad
        maxValue += pad

        Dim rowH As Long = CLng(rect.Height / CDbl(items.Count))
        Dim labelW As Long = CLng(rect.Width * 0.26R)
        Dim plotX As Long = rect.X + labelW
        Dim plotW As Long = CLng(rect.Width * 0.68R)
        Dim dot As Long = 90000L

        For i As System.Int32 = 0 To items.Count - 1
            Dim y As Long = rect.Y + i * rowH + CLng(rowH * 0.48R)
            Dim v1 As System.Double = JsonDouble(items(i), "value", 0.0R)
            Dim v2 As System.Double = JsonDouble(items(i), "benchmark", JsonDouble(items(i), "value2", 0.0R))
            Dim x1 As Long = plotX + CLng(plotW * ((v1 - minValue) / (maxValue - minValue)))
            Dim x2 As Long = plotX + CLng(plotW * ((v2 - minValue) / (maxValue - minValue)))
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X, .Y = y - CLng(rowH * 0.25R), .Width = labelW - CLng(rect.Width * 0.02R), .Height = CLng(rowH * 0.5R)}, JsonString(items(i), "label", JsonString(items(i), "title", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, False)
            AddPrimitiveLine(sp, System.Math.Min(x1, x2), y, System.Math.Max(x1, x2), y, MixHexColors(ctx.Dark, ctx.Background, 0.68R), 1.0R)
            AddPrimitiveShape(sp, New EmuRect With {.X = x2 - dot \ 2, .Y = y - dot \ 2, .Width = dot, .Height = dot}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, ctx.MutedText)
            AddPrimitiveShape(sp, New EmuRect With {.X = x1 - dot \ 2, .Y = y - dot \ 2, .Width = dot, .Height = dot}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, ctx.Accent1)
        Next
    End Sub

    Private Sub RenderDriverTree(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(6).ToList()
        If items.Count = 0 Then Return

        Dim rootTitle As System.String = JsonString(el, "title", JsonString(el, "label", "Key question"))
        Dim rootDetail As System.String = JsonString(el, "detail", "")
        Dim root As New EmuRect With {.X = rect.X, .Y = rect.Y + CLng(rect.Height * 0.29R), .Width = CLng(rect.Width * 0.27R), .Height = CLng(rect.Height * 0.34R)}
        AddPrimitiveShape(sp, root, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, MixHexColors(ctx.Accent1, ctx.Background, 0.84R), ctx.Accent1, 1.0R)
        AddPrimitiveText(sp, New EmuRect With {.X = root.X + CLng(root.Width * 0.08R), .Y = root.Y + CLng(root.Height * 0.15R), .Width = CLng(root.Width * 0.84R), .Height = CLng(root.Height * 0.28R)}, rootTitle, ctx.HeadingFont, ctx.HeadingSize, ctx.Dark, True)
        If Not System.String.IsNullOrWhiteSpace(rootDetail) Then AddPrimitiveText(sp, New EmuRect With {.X = root.X + CLng(root.Width * 0.08R), .Y = root.Y + CLng(root.Height * 0.5R), .Width = CLng(root.Width * 0.84R), .Height = CLng(root.Height * 0.27R)}, rootDetail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)

        Dim rightX As Long = rect.X + CLng(rect.Width * 0.43R)
        Dim rightW As Long = rect.Width - CLng(rect.Width * 0.43R)
        Dim gap As Long = CLng(rect.Height * 0.025R)
        Dim boxH As Long = CLng((rect.Height - gap * (items.Count - 1)) / CDbl(items.Count))
        Dim trunkX As Long = rect.X + CLng(rect.Width * 0.355R)
        Dim rootCenterY As Long = root.Y + root.Height \ 2
        AddPrimitiveLine(sp, root.X + root.Width, rootCenterY, trunkX, rootCenterY, MixHexColors(ctx.Dark, ctx.Background, 0.6R), 1.0R)

        For i As System.Int32 = 0 To items.Count - 1
            Dim y As Long = rect.Y + i * (boxH + gap)
            Dim child As New EmuRect With {.X = rightX, .Y = y, .Width = rightW, .Height = boxH}
            Dim childCenterY As Long = y + boxH \ 2
            AddPrimitiveLine(sp, trunkX, System.Math.Min(rootCenterY, childCenterY), trunkX, System.Math.Max(rootCenterY, childCenterY), MixHexColors(ctx.Dark, ctx.Background, 0.6R), 1.0R)
            AddPrimitiveLine(sp, trunkX, childCenterY, rightX, childCenterY, MixHexColors(ctx.Dark, ctx.Background, 0.6R), 1.0R)
            AddPrimitiveShape(sp, child, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, ctx.Surface, MixHexColors(ctx.Dark, ctx.Background, 0.74R), 0.7R)
            AddPrimitiveText(sp, New EmuRect With {.X = child.X + CLng(child.Width * 0.04R), .Y = child.Y + CLng(child.Height * 0.13R), .Width = CLng(child.Width * 0.92R), .Height = CLng(child.Height * 0.32R)}, JsonString(items(i), "title", JsonString(items(i), "label", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, True)
            Dim detail As System.String = GetComponentDetail(items(i))
            If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = child.X + CLng(child.Width * 0.04R), .Y = child.Y + CLng(child.Height * 0.5R), .Width = CLng(child.Width * 0.92R), .Height = CLng(child.Height * 0.28R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
        Next
    End Sub

    Private Sub RenderAgenda(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(6).ToList()
        If items.Count = 0 Then Return
        Dim rowH As Long = CLng(rect.Height / CDbl(items.Count))
        For i As System.Int32 = 0 To items.Count - 1
            Dim y As Long = rect.Y + i * rowH
            Dim circle As Long = CLng(rowH * 0.48R)
            AddPrimitiveShape(sp, New EmuRect With {.X = rect.X, .Y = y + CLng(rowH * 0.1R), .Width = circle, .Height = circle}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, If(i = 0, ctx.Accent1, MixHexColors(ctx.Accent1, ctx.Light, 0.78R)))
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X, .Y = y + CLng(rowH * 0.18R), .Width = circle, .Height = CLng(rowH * 0.25R)}, (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture), ctx.HeadingFont, 11.0R, If(i = 0, "#FFFFFF", ctx.Accent1), True, DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center)
            AddPrimitiveText(sp, New EmuRect With {.X = rect.X + circle + CLng(rect.Width * 0.03R), .Y = y + CLng(rowH * 0.12R), .Width = rect.Width - circle - CLng(rect.Width * 0.04R), .Height = CLng(rowH * 0.32R)}, JsonString(items(i), "title", JsonString(items(i), "label", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, True)
            Dim detail As System.String = GetComponentDetail(items(i))
            If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = rect.X + circle + CLng(rect.Width * 0.03R), .Y = y + CLng(rowH * 0.5R), .Width = rect.Width - circle - CLng(rect.Width * 0.04R), .Height = CLng(rowH * 0.28R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
        Next
    End Sub

    Private Sub RenderBigNumber(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim value As System.String = JsonString(el, "value", "")
        Dim label As System.String = JsonString(el, "label", "")
        Dim detail As System.String = JsonString(el, "detail", "")
        AddPrimitiveShape(sp, New EmuRect With {.X = rect.X, .Y = rect.Y, .Width = CLng(rect.Width * 0.025R), .Height = rect.Height}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, ctx.Accent1)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.07R), .Y = rect.Y + CLng(rect.Height * 0.08R), .Width = CLng(rect.Width * 0.82R), .Height = CLng(rect.Height * 0.38R)}, value, ctx.HeadingFont, System.Math.Max(34.0R, ctx.HeadingSize * 2.35R), ctx.Accent1, True)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.07R), .Y = rect.Y + CLng(rect.Height * 0.5R), .Width = CLng(rect.Width * 0.82R), .Height = CLng(rect.Height * 0.16R)}, label, ctx.BodyFont, System.Math.Max(14.0R, ctx.BodySize * 1.12R), ctx.Dark, True)
        If Not System.String.IsNullOrWhiteSpace(detail) Then AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.07R), .Y = rect.Y + CLng(rect.Height * 0.7R), .Width = CLng(rect.Width * 0.82R), .Height = CLng(rect.Height * 0.16R)}, detail, ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
    End Sub

    Private Sub RenderQuote(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X, .Y = rect.Y, .Width = CLng(rect.Width * 0.12R), .Height = CLng(rect.Height * 0.3R)}, System.Char.ConvertFromUtf32(&H201C), ctx.HeadingFont, 42.0R, ctx.Accent1, True)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.1R), .Y = rect.Y + CLng(rect.Height * 0.12R), .Width = CLng(rect.Width * 0.82R), .Height = CLng(rect.Height * 0.5R)}, JsonString(el, "quote", JsonString(el, "text", "")), ctx.HeadingFont, System.Math.Max(18.0R, ctx.HeadingSize * 1.2R), ctx.Dark, False)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.1R), .Y = rect.Y + CLng(rect.Height * 0.7R), .Width = CLng(rect.Width * 0.82R), .Height = CLng(rect.Height * 0.14R)}, JsonString(el, "attribution", ""), ctx.BodyFont, ctx.SmallSize, ctx.MutedText, True)
    End Sub

    Private Sub RenderCallout(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        AddPrimitiveShape(sp, rect, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle, MixHexColors(ctx.Accent1, ctx.Light, 0.91R), MixHexColors(ctx.Accent1, ctx.Light, 0.65R), 0.8R)
        AddPrimitiveShape(sp, New EmuRect With {.X = rect.X, .Y = rect.Y, .Width = CLng(rect.Width * 0.018R), .Height = rect.Height}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle, ctx.Accent1)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.06R), .Y = rect.Y + CLng(rect.Height * 0.18R), .Width = CLng(rect.Width * 0.86R), .Height = CLng(rect.Height * 0.24R)}, JsonString(el, "title", JsonString(el, "label", "")), ctx.HeadingFont, ctx.HeadingSize, ctx.Dark, True)
        AddPrimitiveText(sp, New EmuRect With {.X = rect.X + CLng(rect.Width * 0.06R), .Y = rect.Y + CLng(rect.Height * 0.49R), .Width = CLng(rect.Width * 0.86R), .Height = CLng(rect.Height * 0.27R)}, JsonString(el, "detail", JsonString(el, "text", "")), ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
    End Sub

    Private Sub RenderGenericCards(ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart, ByVal el As System.Text.Json.JsonElement, ByVal rect As EmuRect, ByVal ctx As RenderStyleContext)
        Dim items = GetComponentItems(el).Take(6).ToList()
        If items.Count = 0 Then Return
        Dim cols As System.Int32 = If(items.Count <= 3, items.Count, 3)
        Dim rows As System.Int32 = CInt(System.Math.Ceiling(items.Count / CDbl(cols)))
        Dim gapX As Long = CLng(rect.Width * 0.025R)
        Dim gapY As Long = CLng(rect.Height * 0.05R)
        Dim cardW As Long = CLng((rect.Width - gapX * (cols - 1)) / CDbl(cols))
        Dim cardH As Long = CLng((rect.Height - gapY * (rows - 1)) / CDbl(rows))
        For i As System.Int32 = 0 To items.Count - 1
            Dim col As System.Int32 = i Mod cols
            Dim row As System.Int32 = i \ cols
            Dim x As Long = rect.X + col * (cardW + gapX)
            Dim y As Long = rect.Y + row * (cardH + gapY)
            Dim accent As System.String = If(i Mod 3 = 0, ctx.Accent1, If(i Mod 3 = 1, ctx.Accent2, ctx.Accent3))
            AddPrimitiveShape(sp, New EmuRect With {.X = x, .Y = y, .Width = cardW, .Height = cardH}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.RoundRectangle, ctx.Surface, MixHexColors(ctx.Dark, ctx.Light, 0.83R), 0.7R)
            AddPrimitiveShape(sp, New EmuRect With {.X = x + CLng(cardW * 0.06R), .Y = y + CLng(cardH * 0.11R), .Width = CLng(cardW * 0.07R), .Height = CLng(cardH * 0.07R)}, DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Ellipse, accent)
            AddPrimitiveText(sp, New EmuRect With {.X = x + CLng(cardW * 0.06R), .Y = y + CLng(cardH * 0.27R), .Width = CLng(cardW * 0.87R), .Height = CLng(cardH * 0.2R)}, JsonString(items(i), "title", JsonString(items(i), "label", "")), ctx.BodyFont, ctx.BodySize, ctx.Dark, True)
            AddPrimitiveText(sp, New EmuRect With {.X = x + CLng(cardW * 0.06R), .Y = y + CLng(cardH * 0.52R), .Width = CLng(cardW * 0.87R), .Height = CLng(cardH * 0.3R)}, GetComponentDetail(items(i)), ctx.BodyFont, ctx.SmallSize, ctx.MutedText, False)
        Next
    End Sub

    Private Sub AddBuiltinIcon(
        ByVal presPart As DocumentFormat.OpenXml.Packaging.PresentationPart,
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement)

        Dim name As System.String = JsonString(el, "name", "check")
        Dim ctx As RenderStyleContext = GetRenderStyleContext(sp)
        Dim color As System.String = JsonString(el, "color", ctx.Accent1)
        Dim tf As System.Text.Json.JsonElement
        If Not el.TryGetProperty("transform", tf) Then Return
        AddSvgIconFromString(sp, GetBuiltinIconSvg(name, color), GetTransformFromJson(presPart, tf), "Icon " & name)
    End Sub

    Private Sub AddSvgIconFromString(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal svg As System.String,
        ByVal xfrm As DocumentFormat.OpenXml.Drawing.Transform2D,
        ByVal name As System.String)

        Dim id As UInteger = NextShapeId(sp)
        Dim svgPart As DocumentFormat.OpenXml.Packaging.ImagePart = sp.AddImagePart(DocumentFormat.OpenXml.Packaging.ImagePartType.Svg)
        Using ms As New System.IO.MemoryStream(System.Text.Encoding.UTF8.GetBytes(svg))
            svgPart.FeedData(ms)
        End Using
        Dim relId As System.String = sp.GetIdOfPart(svgPart)
        Dim nvPic As New DocumentFormat.OpenXml.Presentation.NonVisualPictureProperties(
            New DocumentFormat.OpenXml.Presentation.NonVisualDrawingProperties() With {.Id = id, .Name = name},
            New DocumentFormat.OpenXml.Presentation.NonVisualPictureDrawingProperties(New DocumentFormat.OpenXml.Drawing.PictureLocks() With {.NoChangeAspect = True}),
            New DocumentFormat.OpenXml.Presentation.ApplicationNonVisualDrawingProperties())
        Dim blipFill As New DocumentFormat.OpenXml.Presentation.BlipFill(
            New DocumentFormat.OpenXml.Drawing.Blip() With {.Embed = relId, .CompressionState = DocumentFormat.OpenXml.Drawing.BlipCompressionValues.Print},
            New DocumentFormat.OpenXml.Drawing.Stretch(New DocumentFormat.OpenXml.Drawing.FillRectangle()))
        Dim spPr As New DocumentFormat.OpenXml.Presentation.ShapeProperties()
        spPr.Append(xfrm)
        spPr.Append(New DocumentFormat.OpenXml.Drawing.PresetGeometry(New DocumentFormat.OpenXml.Drawing.AdjustValueList()) With {.Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle})
        sp.Slide.CommonSlideData.ShapeTree.Append(New DocumentFormat.OpenXml.Presentation.Picture(nvPic, blipFill, spPr))
    End Sub

    Private Function GetBuiltinIconSvg(ByVal name As System.String, ByVal color As System.String) As System.String
        Dim c As System.String = NormalizeHexColor(color, "#2563EB")
        Dim body As System.String
        Select Case If(name, "check").Trim().ToLowerInvariant()
            Case "arrow", "next"
                body = "<path d='M5 12h14M13 6l6 6-6 6'/>"
            Case "growth", "trend"
                body = "<path d='M4 17l5-5 4 3 7-8'/><path d='M15 7h5v5'/>"
            Case "target"
                body = "<circle cx='12' cy='12' r='8'/><circle cx='12' cy='12' r='4'/><circle cx='12' cy='12' r='1' fill='" & c & "' stroke='none'/>"
            Case "people", "team"
                body = "<circle cx='9' cy='8' r='3'/><circle cx='17' cy='9' r='2.5'/><path d='M3 19c0-4 3-6 6-6s6 2 6 6M14 14c3 0 6 2 6 5'/>"
            Case "shield", "security"
                body = "<path d='M12 3l7 3v5c0 5-3 8-7 10-4-2-7-5-7-10V6l7-3z'/><path d='M9 12l2 2 4-4'/>"
            Case "clock", "time"
                body = "<circle cx='12' cy='12' r='9'/><path d='M12 7v6l4 2'/>"
            Case "lightbulb", "idea"
                body = "<path d='M9 18h6M10 21h4'/><path d='M8 14c-2-2-2-6 0-8 2-2 6-2 8 0 2 2 2 6 0 8-1 1-2 2-2 4h-4c0-2-1-3-2-4z'/>"
            Case "warning", "risk"
                body = "<path d='M12 3l10 18H2L12 3z'/><path d='M12 9v5M12 18h.01'/>"
            Case "star"
                body = "<path d='M12 3l2.7 5.5 6.1.9-4.4 4.3 1 6.1-5.4-2.9-5.4 2.9 1-6.1-4.4-4.3 6.1-.9L12 3z'/>"
            Case "chart", "data"
                body = "<path d='M4 20V10h4v10M10 20V4h4v16M16 20v-7h4v7M3 20h18'/>"
            Case "globe"
                body = "<circle cx='12' cy='12' r='9'/><path d='M3 12h18M12 3c3 3 4 6 4 9s-1 6-4 9M12 3c-3 3-4 6-4 9s1 6 4 9'/>"
            Case "document"
                body = "<path d='M6 3h8l4 4v14H6z'/><path d='M14 3v5h5M9 12h6M9 16h6'/>"
            Case "gear", "settings"
                body = "<circle cx='12' cy='12' r='3'/><path d='M12 2v3M12 19v3M2 12h3M19 12h3M5 5l2 2M17 17l2 2M19 5l-2 2M7 17l-2 2'/>"
            Case Else
                body = "<path d='M5 12l4 4L19 6'/>"
        End Select
        Return "<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24' fill='none' stroke='" & c & "' stroke-width='2' stroke-linecap='round' stroke-linejoin='round'>" & body & "</svg>"
    End Function

#End Region

    Private Sub SetTextWithPlaceholder(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal placeholderEl As System.Text.Json.JsonElement,
        ByVal text As System.String,
        ByVal el As System.Text.Json.JsonElement)

        Dim targetShape As DocumentFormat.OpenXml.Presentation.Shape = FindShapeByPlaceholderElement(sp, placeholderEl)
        If targetShape Is Nothing Then Return
        SetShapeSingleTextPreserveStyle(targetShape, text, el, forceNoBullet:=True)
        ApplyGeneratedPlaceholderTextFit(sp, targetShape, text)
        sp.Slide.Save()
    End Sub

    Private Sub SetBulletsWithPlaceholder(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal el As System.Text.Json.JsonElement)

        Dim targetShape As DocumentFormat.OpenXml.Presentation.Shape = Nothing
        Dim placeholderEl As System.Text.Json.JsonElement
        If el.TryGetProperty("placeholder", placeholderEl) Then
            targetShape = FindShapeByPlaceholderElement(sp, placeholderEl)
        Else
            For Each shp As DocumentFormat.OpenXml.Presentation.Shape In sp.Slide.CommonSlideData.ShapeTree.Elements(Of DocumentFormat.OpenXml.Presentation.Shape)()
                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                If IsBodyLikePlaceholder(ph) Then
                    targetShape = shp
                    Exit For
                End If
            Next
        End If
        If targetShape Is Nothing Then Return
        SetShapeBulletsPreserveStyle(targetShape, el)
        sp.Slide.Save()
    End Sub

    Private Function FindShapeByPlaceholderElement(
        ByVal sp As DocumentFormat.OpenXml.Packaging.SlidePart,
        ByVal placeholderEl As System.Text.Json.JsonElement) As DocumentFormat.OpenXml.Presentation.Shape

        If sp?.Slide?.CommonSlideData?.ShapeTree Is Nothing Then Return Nothing
        Dim allShapes = sp.Slide.CommonSlideData.ShapeTree.Descendants(Of DocumentFormat.OpenXml.Presentation.Shape)()

        If placeholderEl.ValueKind = System.Text.Json.JsonValueKind.Object Then
            Dim tmp As System.Text.Json.JsonElement

            If placeholderEl.TryGetProperty("shapeId", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then
                Dim requestedId As UInteger
                If tmp.TryGetUInt32(requestedId) Then
                    Dim byId = allShapes.FirstOrDefault(Function(shp) shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id IsNot Nothing AndAlso shp.NonVisualShapeProperties.NonVisualDrawingProperties.Id.Value = requestedId)
                    If byId IsNot Nothing Then Return byId
                End If
            End If

            If placeholderEl.TryGetProperty("name", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then
                Dim requestedName As System.String = tmp.GetString()
                Dim byName = allShapes.FirstOrDefault(Function(shp)
                                                          Dim n As System.String = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
                                                          Return System.String.Equals(n, requestedName, System.StringComparison.OrdinalIgnoreCase) OrElse n.IndexOf(requestedName, System.StringComparison.OrdinalIgnoreCase) >= 0
                                                      End Function)
                If byName IsNot Nothing Then Return byName
            End If

            Dim requestedIndex As Nullable(Of UInteger) = Nothing
            If placeholderEl.TryGetProperty("index", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.Number Then
                Dim idx As UInteger
                If tmp.TryGetUInt32(idx) Then requestedIndex = idx
            End If
            If requestedIndex.HasValue Then
                Dim byIndex = allShapes.FirstOrDefault(Function(shp)
                                                           Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                                           Return ph?.Index IsNot Nothing AndAlso ph.Index.Value = requestedIndex.Value
                                                       End Function)
                If byIndex IsNot Nothing Then Return byIndex
            End If

            Dim requestedSemanticRole As System.String = System.String.Empty
            If placeholderEl.TryGetProperty("semanticRole", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then requestedSemanticRole = tmp.GetString()
            If Not System.String.IsNullOrWhiteSpace(requestedSemanticRole) Then
                Dim semanticMatches As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
                    allShapes.
                        Where(
                            Function(shp)
                                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                If ph Is Nothing Then Return False
                                Dim semanticRole As System.String = ResolveEffectivePlaceholderSemanticRole(sp, shp)
                                Return System.String.Equals(semanticRole, requestedSemanticRole, System.StringComparison.OrdinalIgnoreCase)
                            End Function).
                        ToList()

                If semanticMatches.Count > 0 Then
                    If System.String.Equals(requestedSemanticRole, "body", System.StringComparison.OrdinalIgnoreCase) Then
                        Return semanticMatches.
                            OrderByDescending(
                                Function(shp)
                                    Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shp)
                                    If xfrm?.Extents Is Nothing Then Return 0.0R
                                    Return CDbl(xfrm.Extents.Cx.Value) * CDbl(xfrm.Extents.Cy.Value)
                                End Function).
                            FirstOrDefault()
                    End If
                    Return semanticMatches.FirstOrDefault()
                End If

                ' A specific semantic target such as location/subtitle/logo must never silently fall
                ' through to an arbitrary body placeholder. Misplacing metadata is worse than omitting it.
                If Not System.String.Equals(requestedSemanticRole, "body", System.StringComparison.OrdinalIgnoreCase) Then
                    Return Nothing
                End If
            End If

            Dim requestedType As System.String = System.String.Empty
            If placeholderEl.TryGetProperty("type", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then requestedType = tmp.GetString()
            Dim requestedRole As System.String = System.String.Empty
            If placeholderEl.TryGetProperty("role", tmp) AndAlso tmp.ValueKind = System.Text.Json.JsonValueKind.String Then requestedRole = tmp.GetString()

            If System.String.Equals(requestedType, "body", System.StringComparison.OrdinalIgnoreCase) AndAlso
               System.String.IsNullOrWhiteSpace(requestedRole) Then
                requestedRole = "body"
            End If

            If Not System.String.IsNullOrWhiteSpace(requestedType) Then
                For Each shp As DocumentFormat.OpenXml.Presentation.Shape In allShapes
                    Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph Is Nothing Then Continue For
                    Select Case requestedType.Trim().ToLowerInvariant()
                        Case "body"
                            ' Do not immediately return the first p:ph type="body": in many
                            ' corporate templates that is only an eyebrow strip. The requested
                            ' role branch below will choose the largest semantic body instead.
                            Continue For
                        Case "object"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Object Then Return shp
                        Case "subtitle"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.SubTitle Then Return shp
                            If ph.Type Is Nothing AndAlso ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then Return shp
                        Case "title"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.Title Then Return shp
                        Case "centeredtitle"
                            If ph.Type IsNot Nothing AndAlso ph.Type.Value = DocumentFormat.OpenXml.Presentation.PlaceholderValues.CenteredTitle Then Return shp
                        Case Else
                            If ph.Type IsNot Nothing AndAlso System.String.Equals(ph.Type.Value.ToString(), requestedType, System.StringComparison.OrdinalIgnoreCase) Then Return shp
                    End Select
                Next
            End If

            If Not System.String.IsNullOrWhiteSpace(requestedRole) Then
                Dim roleMatches As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
                    allShapes.
                        Where(
                            Function(shp)
                                Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                                If ph Is Nothing Then Return False
                                Dim role As System.String = ResolvePlaceholderRoleForJson(ph, If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty))
                                If Not System.String.Equals(role, requestedRole, System.StringComparison.OrdinalIgnoreCase) Then Return False
                                If System.String.Equals(requestedRole, "body", System.StringComparison.OrdinalIgnoreCase) Then
                                    Return System.String.Equals(ResolveEffectivePlaceholderSemanticRole(sp, shp), "body", System.StringComparison.OrdinalIgnoreCase)
                                End If
                                Return True
                            End Function).
                        ToList()

                If roleMatches.Count > 0 Then
                    If System.String.Equals(requestedRole, "body", System.StringComparison.OrdinalIgnoreCase) Then
                        Return roleMatches.
                            OrderByDescending(
                                Function(shp)
                                    Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shp)
                                    If xfrm?.Extents Is Nothing Then Return 0.0R
                                    Return CDbl(xfrm.Extents.Cx.Value) * CDbl(xfrm.Extents.Cy.Value)
                                End Function).
                            FirstOrDefault()
                    End If
                    Return roleMatches.FirstOrDefault()
                End If
            End If

        ElseIf placeholderEl.ValueKind = System.Text.Json.JsonValueKind.String Then
            Dim nameToFind As System.String = placeholderEl.GetString()
            For Each shp As DocumentFormat.OpenXml.Presentation.Shape In allShapes
                Dim nm As System.String = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, System.String.Empty)
                If nm.IndexOf(nameToFind, System.StringComparison.OrdinalIgnoreCase) >= 0 Then Return shp
            Next
        End If

        Dim fallbackBodies As System.Collections.Generic.List(Of DocumentFormat.OpenXml.Presentation.Shape) =
            allShapes.
                Where(
                    Function(shp)
                        Dim ph As DocumentFormat.OpenXml.Presentation.PlaceholderShape = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                        If ph Is Nothing OrElse Not IsBodyLikePlaceholder(ph) Then Return False
                        Return System.String.Equals(ResolveEffectivePlaceholderSemanticRole(sp, shp), "body", System.StringComparison.OrdinalIgnoreCase)
                    End Function).
                ToList()

        If fallbackBodies.Count > 0 Then
            Return fallbackBodies.
                OrderByDescending(
                    Function(shp)
                        Dim xfrm As DocumentFormat.OpenXml.Drawing.Transform2D = GetEffectiveTransformForShape(sp, shp)
                        If xfrm?.Extents Is Nothing Then Return 0.0R
                        Return CDbl(xfrm.Extents.Cx.Value) * CDbl(xfrm.Extents.Cy.Value)
                    End Function).
                FirstOrDefault()
        End If
        Return Nothing
    End Function


End Class
