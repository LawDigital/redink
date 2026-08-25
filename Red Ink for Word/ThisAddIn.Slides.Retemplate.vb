' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Slides.Retemplate.vb
' Purpose: Re-skins an existing .pptx by applying the design (master / layouts /
'          theme / fonts / colors / custom assets) of a second .pptx, without
'          losing any content of the source deck.
'
' Strategy:
'   1. Copy the template .pptx to the output path and strip its sample slides.
'   2. For each source slide pick the best matching template layout
'      (heuristic, optional LLM, optional user review dialog).
'   3. Create a new slide in the output that references the chosen layout.
'   4. Reflow recognized placeholder content (Title / Subtitle / Body) into
'      the template's placeholders.
'   5. Deep-copy every non-placeholder element of the source slide verbatim
'      (pictures, tables, charts, SmartArt, media, groups, connectors, OLE,
'      ink, SVG/PNG fallbacks inside mc:AlternateContent, custom XML, …).
'   6. Copy notes, transitions, timing/animations (with shape-id remap),
'      per-slide comments, hidden/show flags.
'   7. Optionally rescale source shapes if slide size differs.
'   8. Preserve presentation-level extras (embedded fonts, comment authors,
'      custom XML, defaultTextStyle if missing).
'
' External dependencies: DocumentFormat.OpenXml, System.Text.Json,
'   SharedLibrary.SharedMethods (UI + LLM), NetOffice.PowerPointApi
'   (only for optional thumbnail preview and PNG raster fallback).
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.IO
Imports System.Text.Json
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Wordprocessing
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports A = DocumentFormat.OpenXml.Drawing
Imports P = DocumentFormat.OpenXml.Presentation
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

#Region "Public entry point"

    ''' <summary>
    ''' UI entry point: asks the user for source / template / output / options,
    ''' then executes the re-template pipeline.
    ''' </summary>
    Public Async Sub RetemplatePresentation_UI()

        If INILoadFail() Then Return

        Try
            ' ---------------- Pick files ----------------
            Dim srcPath As String = PickPptx("Select the source presentation (content)")
            If String.IsNullOrWhiteSpace(srcPath) Then Return

            Dim tplPath As String = PickPptx("Select the template presentation (design)")
            If String.IsNullOrWhiteSpace(tplPath) Then Return

            Dim outPath As String = PickPptxSave(srcPath)
            If String.IsNullOrWhiteSpace(outPath) Then Return

            ' ---------------- Parameter menu ----------------
            Dim defUseLlm As Boolean = True
            Dim defReview As Boolean = True
            Dim defPreserveTypo As Boolean = False
            Dim defScale As String = "Keep template size (warn on overflow)"
            Dim defCopyFonts As Boolean = True
            Dim defCopySrcFonts As Boolean = False
            Dim defDeepChart As Boolean = True

            Dim hasSecondary As Boolean =
                (Not String.IsNullOrWhiteSpace(INI_AlternateModelPath)) OrElse INI_SecondAPI

            Dim pUseLlm As New SLib.InputParameter("Use LLM to assist layout matching", defUseLlm)
            Dim pReview As New SLib.InputParameter("Show layout review dialog before conversion", defReview)
            Dim pPreserveTypo As New SLib.InputParameter("Preserve source typography (size/color/font)", defPreserveTypo)
            Dim pScale As New SLib.InputParameter("If slide size differs", defScale)
            pScale.Options = New List(Of String) From {
                "Keep template size (warn on overflow)",
                "Scale source shapes to fit template"
            }
            Dim pCopyFonts As New SLib.InputParameter("Copy embedded fonts from template", defCopyFonts)
            Dim pCopySrcFonts As New SLib.InputParameter("Also copy embedded fonts from source", defCopySrcFonts)
            Dim pDeepChart As New SLib.InputParameter("Deep-copy charts/SmartArt (PNG fallback on error)", defDeepChart)
            Dim pSecond As SLib.InputParameter = Nothing
            If hasSecondary Then
                pSecond = New SLib.InputParameter("Use secondary model (for layout matching)", False)
            End If

            Dim params() As SLib.InputParameter =
                If(hasSecondary,
                   New SLib.InputParameter() {pUseLlm, pReview, pPreserveTypo, pScale, pCopyFonts, pCopySrcFonts, pDeepChart, pSecond},
                   New SLib.InputParameter() {pUseLlm, pReview, pPreserveTypo, pScale, pCopyFonts, pCopySrcFonts, pDeepChart})

            If ShowCustomVariableInputForm(
                "Apply the design of the template to the source presentation." & vbCrLf &
                "Source: " & System.IO.Path.GetFileName(srcPath) & vbCrLf &
                "Template: " & System.IO.Path.GetFileName(tplPath) & vbCrLf &
                "Output: " & System.IO.Path.GetFileName(outPath),
                AN & " Re-template Presentation", params) = False Then Return

            Dim opts As New RetemplateOptions With {
                .UseLlm = CBool(params(0).Value),
                .ShowReviewDialog = CBool(params(1).Value),
                .PreserveSourceTypography = CBool(params(2).Value),
                .ScaleToTemplateSize = (CStr(params(3).Value).IndexOf("Scale", StringComparison.OrdinalIgnoreCase) >= 0),
                .CopyTemplateEmbeddedFonts = CBool(params(4).Value),
                .CopySourceEmbeddedFonts = CBool(params(5).Value),
                .DeepCopyChartSmartArt = CBool(params(6).Value),
                .UseSecondaryApi = If(hasSecondary, CBool(params(7).Value), False)
            }

            ' ---------------- Run pipeline ----------------
            Dim result As RetemplateResult = Await RetemplatePresentation(srcPath, tplPath, outPath, opts)

            RunOnUiThread(
                Sub()
                    If result.Success Then
                        Dim msg As String = "Re-template completed." & vbCrLf &
                                            "Output: " & outPath & vbCrLf &
                                            "Slides processed: " & result.SlideMappings.Count
                        If result.Warnings.Count > 0 Then
                            msg &= vbCrLf & vbCrLf & "Warnings:" & vbCrLf &
                                   String.Join(vbCrLf, result.Warnings.Take(20))
                            If result.Warnings.Count > 20 Then
                                msg &= vbCrLf & "… (" & (result.Warnings.Count - 20) & " more)"
                            End If
                        End If

                        If ShowCustomYesNoBox(msg & vbCrLf & vbCrLf & "Open the result now?", "Open", "Close") = 1 Then
                            Try : System.Diagnostics.Process.Start(outPath) : Catch : End Try
                        End If
                    Else
                        ShowCustomMessageBox("Re-template failed." & vbCrLf &
                                             If(result.Errors.Count > 0, String.Join(vbCrLf, result.Errors), ""))
                    End If
                End Sub)

        Catch ex As Exception
            RunOnUiThread(
                Sub()
                    ShowCustomMessageBox("Unexpected error: " & ex.Message)
                End Sub)
        End Try
    End Sub

    ''' <summary>
    ''' Programmatic entry point: applies the design of <paramref name="templatePath"/>
    ''' to the content of <paramref name="sourcePath"/>, writing to <paramref name="outputPath"/>.
    ''' Shows the standard progress bar until the review dialog appears (or until completion).
    ''' </summary>
    Public Async Function RetemplatePresentation(
            sourcePath As String,
            templatePath As String,
            outputPath As String,
            Optional options As RetemplateOptions = Nothing) As Task(Of RetemplateResult)

        Dim res As New RetemplateResult
        If options Is Nothing Then options = New RetemplateOptions()

        ' -------- Basic validation --------
        If Not System.IO.File.Exists(sourcePath) Then
            res.Errors.Add("Source file not found: " & sourcePath) : Return res
        End If
        If Not System.IO.File.Exists(templatePath) Then
            res.Errors.Add("Template file not found: " & templatePath) : Return res
        End If
        If Not IsValidPptxPackage(sourcePath) Then
            res.Errors.Add("Source file is not a valid .pptx package.") : Return res
        End If
        If Not IsValidPptxPackage(templatePath) Then
            res.Errors.Add("Template file is not a valid .pptx package.") : Return res
        End If

        ' ============================================================
        ' Standard progress bar via ProgressBarModule
        ' ============================================================
        ProgressBarModule.CancelOperation = False
        ProgressBarModule.GlobalProgressMax = 8
        ProgressBarModule.GlobalProgressValue = 0
        ProgressBarModule.GlobalProgressLabel = "Preparing…"
        ShowProgressBarInSeparateThread(AN & " Re-template Presentation", "Preparing…")

        Try
            ' -------- 1) Build output package from template --------
            ProgressBarModule.GlobalProgressValue = 1
            ProgressBarModule.GlobalProgressLabel = "Preparing template…"
            Try
                System.IO.File.Copy(templatePath, outputPath, True)
            Catch ex As Exception
                res.Errors.Add("Could not create output file: " & ex.Message) : Return res
            End Try

            Try
                StripSampleSlidesFromPackage(outputPath)
            Catch ex As Exception
                res.Errors.Add("Could not strip template sample slides: " & ex.Message) : Return res
            End Try

            ' -------- 2) Build layout catalog from template --------
            ProgressBarModule.GlobalProgressValue = 2
            ProgressBarModule.GlobalProgressLabel = "Analyzing template layouts…"
            Dim catalog As List(Of LayoutSignature) = BuildLayoutCatalog(outputPath)
            If catalog.Count = 0 Then
                res.Errors.Add("Template has no slide layouts.") : Return res
            End If

            ' -------- 3) Build source slide signatures --------
            ProgressBarModule.GlobalProgressValue = 3
            ProgressBarModule.GlobalProgressLabel = "Analyzing source slides…"
            Dim srcSignatures As List(Of SlideSignature) = BuildSlideSignatures(sourcePath)
            If srcSignatures.Count = 0 Then
                res.Warnings.Add("Source has no slides; output contains only the template masters.")
                res.Success = True : Return res
            End If

            ' Recalculate max: 5 prep stages + N slides + 2 finalize stages
            ProgressBarModule.GlobalProgressMax = 5 + srcSignatures.Count + 2

            ' -------- 4) Match layouts --------
            ProgressBarModule.GlobalProgressValue = 4
            ProgressBarModule.GlobalProgressLabel = "Matching layouts (heuristic)…"
            Dim mappings As New List(Of SlideLayoutMapping)
            For Each sig In srcSignatures
                Dim top3 = ScoreLayoutCandidates(sig, catalog).Take(3).ToList()
                Dim choice As LayoutSignature = If(top3.Count > 0, top3(0), catalog(0))
                mappings.Add(New SlideLayoutMapping With {
                    .SourceIndex = sig.Index,
                    .SourceKey = sig.Key,
                    .SourceTitle = sig.Title,
                    .LayoutRelId = choice.LayoutRelId,
                    .LayoutName = choice.Name,
                    .Confidence = If(top3.Count > 0, top3(0).LastScore, 0.0),
                    .Rationale = "heuristic"
                })
            Next

            ' -------- 4b) Optional LLM pass --------
            If options.UseLlm Then
                ProgressBarModule.GlobalProgressValue = 5
                ProgressBarModule.GlobalProgressLabel = "Refining layout matches (LLM)…"
                Try
                    Await RefineMappingsWithLlm(srcSignatures, catalog, mappings, options)
                Catch ex As Exception
                    res.Warnings.Add("LLM layout matching failed, keeping heuristic: " & ex.Message)
                End Try
            End If

            ' -------- 4c) Optional user review --------
            If options.ShowReviewDialog Then
                ' Dismiss progress bar BEFORE showing the modal review dialog
                ProgressBarModule.CancelOperation = True

                Dim reviewResult As Boolean =
                    RunOnUiThread(Function() ShowLayoutReviewDialog(mappings, catalog, sourcePath))

                If Not reviewResult Then
                    res.Errors.Add("User cancelled layout review.") : Return res
                End If

                ' Restart progress bar for the conversion phase
                ProgressBarModule.CancelOperation = False
                ProgressBarModule.GlobalProgressMax = srcSignatures.Count + 2
                ProgressBarModule.GlobalProgressValue = 0
                ProgressBarModule.GlobalProgressLabel = "Converting slides…"
                ShowProgressBarInSeparateThread(AN & " Re-template Presentation", "Converting slides…")
            End If

            ' -------- 5) Execute conversion --------
            Dim currentStage As String = "initialization"
            Dim srcBytes As Byte() = Nothing
            Dim srcMs As MemoryStream = Nothing
            Dim srcDoc As PresentationDocument = Nothing
            Dim outDoc As PresentationDocument = Nothing
            Dim tempClonePath As String = Nothing

            Try
                currentStage = "loading source file into memory"
                srcBytes = System.IO.File.ReadAllBytes(sourcePath)
                _rasterSourcePath = sourcePath

                currentStage = "opening source package (read-only)"
                srcMs = New MemoryStream(srcBytes, writable:=False)
                srcDoc = PresentationDocument.Open(srcMs, False)

                currentStage = "preparing file-backed clone of the template"
                tempClonePath = System.IO.Path.Combine(
                    System.IO.Path.GetTempPath(),
                    "RedInk_Retemplate_" & Guid.NewGuid().ToString("N") & ".pptx")
                System.IO.File.Copy(outputPath, tempClonePath, overwrite:=True)

                currentStage = "opening clone package (file-backed, editable)"
                Dim outOpen As New OpenSettings With {.AutoSave = False}
                outDoc = PresentationDocument.Open(tempClonePath, True, outOpen)

                Dim srcPres = srcDoc.PresentationPart
                Dim outPres = outDoc.PresentationPart

                currentStage = "handling slide size"
                HandleSlideSize(srcPres, outPres, options, res)

                currentStage = "copying embedded fonts"
                If options.CopySourceEmbeddedFonts Then
                    Try
                        CopyEmbeddedFonts(srcPres, outPres)
                    Catch ex As Exception
                        res.Warnings.Add("Embedded fonts (source) could not be copied: " & ex.Message)
                    End Try
                End If

                currentStage = "copying presentation-level custom XML"
                Try
                    CopyCustomXmlParts(srcPres, outPres)
                Catch ex As Exception
                    res.Warnings.Add("Presentation-level custom XML parts skipped: " & ex.Message)
                End Try

                currentStage = "ensuring comment authors"
                EnsureCommentAuthors(srcPres, outPres)

                currentStage = "enumerating source slides"
                Dim srcIds = srcPres.Presentation.SlideIdList.
                             Elements(Of P.SlideId)().
                             ToList()
                For i As Integer = 0 To srcIds.Count - 1
                    If ProgressBarModule.CancelOperation Then
                        res.Errors.Add("Operation cancelled by user.")
                        Return res
                    End If

                    Dim slideLabel As String = "Slide " & (i + 1) & " of " & srcIds.Count
                    If Not String.IsNullOrWhiteSpace(mappings(i).SourceTitle) Then
                        slideLabel &= " – " & mappings(i).SourceTitle
                    End If
                    ProgressBarModule.GlobalProgressValue = i + 1
                    ProgressBarModule.GlobalProgressLabel = slideLabel

                    Dim map = mappings(i)
                    Dim srcSid = srcIds(i)
                    Dim srcSp = TryCast(srcPres.GetPartById(srcSid.RelationshipId), SlidePart)
                    If srcSp Is Nothing Then
                        res.Warnings.Add("Source slide " & (i + 1) & " could not be read; skipped.")
                        Continue For
                    End If

                    currentStage = "processing slide " & (i + 1).ToString()
                    Try
                        ProcessSourceSlide(srcPres, outPres, srcSp, srcSid, map, options, res)
                    Catch ex As Exception
                        res.Warnings.Add("Slide " & (i + 1) & " ('" & map.SourceTitle & "') failed: " & ex.Message)
                    End Try
                Next

                currentStage = "writing final package via manual ZIP writer"
                ProgressBarModule.GlobalProgressValue = srcIds.Count + 1
                ProgressBarModule.GlobalProgressLabel = "Writing output file…"
                WriteOpenXmlPackageAsZip(outDoc, tempClonePath, outputPath, res)

                currentStage = "disposing output (polluted) package"
                Try
                    outDoc.Dispose()
                Catch
                End Try
                outDoc = Nothing

                currentStage = "disposing source package"
                Try
                    srcDoc.Dispose()
                Catch
                End Try
                srcDoc = Nothing

                If srcMs IsNot Nothing Then
                    srcMs.Dispose()
                    srcMs = Nothing
                End If

                currentStage = "validating output package"
                ProgressBarModule.GlobalProgressValue = srcIds.Count + 2
                ProgressBarModule.GlobalProgressLabel = "Validating…"
                Try
                    Dim vErr = ValidatePptx(outputPath)
                    If Not String.IsNullOrEmpty(vErr) Then
                        res.Warnings.Add("OpenXML validation: " & vErr)
                    End If
                Catch ex As Exception
                    res.Warnings.Add("Validation skipped: " & ex.Message)
                End Try

                res.Success = True
                res.SlideMappings = mappings
                Return res

            Catch ex As OpenXmlPackageException
                res.Errors.Add("Stage: " & currentStage & vbCrLf & ex.ToString())
                Return res

            Catch ex As Exception
                res.Errors.Add("Stage: " & currentStage & vbCrLf & ex.ToString())
                Return res

            Finally
                Try
                    If outDoc IsNot Nothing Then outDoc.Dispose()
                Catch
                End Try
                Try
                    If srcDoc IsNot Nothing Then srcDoc.Dispose()
                Catch
                End Try
                Try
                    If srcMs IsNot Nothing Then srcMs.Dispose()
                Catch
                End Try
                Try
                    If tempClonePath IsNot Nothing AndAlso
                       System.IO.File.Exists(tempClonePath) Then
                        System.IO.File.Delete(tempClonePath)
                    End If
                Catch
                End Try
            End Try

        Finally
            ' Always dismiss progress bar
            ProgressBarModule.CancelOperation = True
        End Try
    End Function





    Private Sub RunOnUiThread(action As System.Action)
        If action Is Nothing Then Return

        If mainThreadControl Is Nothing OrElse
           mainThreadControl.IsDisposed OrElse
           Not mainThreadControl.IsHandleCreated Then
            action.Invoke()
            Return
        End If

        If mainThreadControl.InvokeRequired Then
            mainThreadControl.Invoke(action)
        Else
            action.Invoke()
        End If
    End Sub

    Private Function RunOnUiThread(Of T)(func As Func(Of T)) As T
        If func Is Nothing Then Return Nothing

        If mainThreadControl Is Nothing OrElse
           mainThreadControl.IsDisposed OrElse
           Not mainThreadControl.IsHandleCreated Then
            Return func.Invoke()
        End If

        If mainThreadControl.InvokeRequired Then
            Return CType(mainThreadControl.Invoke(func), T)
        Else
            Return func.Invoke()
        End If
    End Function



#End Region

#Region "DTOs"

    Public Class RetemplateOptions
        Public Property UseLlm As Boolean = True
        Public Property ShowReviewDialog As Boolean = True
        Public Property PreserveSourceTypography As Boolean = False
        Public Property ScaleToTemplateSize As Boolean = False
        Public Property CopyTemplateEmbeddedFonts As Boolean = True
        Public Property CopySourceEmbeddedFonts As Boolean = False
        Public Property DeepCopyChartSmartArt As Boolean = True
        Public Property UseSecondaryApi As Boolean = False
    End Class

    Public Class RetemplateResult
        Public Property Success As Boolean = False
        Public Property SlideMappings As List(Of SlideLayoutMapping) = New List(Of SlideLayoutMapping)()
        Public Property Warnings As List(Of String) = New List(Of String)()
        Public Property Errors As List(Of String) = New List(Of String)()
    End Class

    Public Class SlideLayoutMapping
        Public Property SourceIndex As Integer
        Public Property SourceKey As String
        Public Property SourceTitle As String
        Public Property LayoutRelId As String
        Public Property LayoutName As String
        Public Property Confidence As Double
        Public Property Rationale As String
    End Class

    Public Class SlideSignature
        Public Property Index As Integer
        Public Property Key As String
        Public Property Title As String
        Public Property HasTitle As Boolean
        Public Property HasSubTitle As Boolean
        Public Property BodyCount As Integer
        Public Property PictureCount As Integer
        Public Property TableCount As Integer
        Public Property ChartCount As Integer
        Public Property SmartArtCount As Integer
        Public Property MediaCount As Integer
        Public Property ShapeCount As Integer
        Public Property IsCover As Boolean
        Public Property IsSectionHeader As Boolean
        Public Property IsBlank As Boolean
        Public Property TwoColumnHint As Boolean

        ''' <summary>
        ''' The name of the layout used by the source slide (e.g. "Title and Content").
        ''' Gives the LLM a strong semantic hint about the original intent.
        ''' </summary>
        Public Property SourceLayoutName As String

        ''' <summary>
        ''' First ~200 characters of the main body text content for LLM context.
        ''' Helps distinguish section headers (short) from content slides (long).
        ''' </summary>
        Public Property ContentPreview As String
    End Class

    Public Class LayoutSignature
        Public Property LayoutRelId As String
        Public Property Name As String
        Public Property Uri As String
        Public Property HasTitle As Boolean
        Public Property HasCenteredTitle As Boolean
        Public Property HasSubTitle As Boolean
        Public Property BodyCount As Integer
        Public Property PicturePlaceholderCount As Integer
        Public Property IsCoverLike As Boolean
        Public Property IsSectionHeaderLike As Boolean
        Public Property IsBlank As Boolean
        Public Property IsTwoContent As Boolean
        <System.Text.Json.Serialization.JsonIgnore> Public Property LastScore As Double
    End Class

#End Region

#Region "File picker helpers"

    Private Function PickPptx(title As String) As String
        Using dlg As New OpenFileDialog()
            dlg.Title = title
            dlg.Filter = "PowerPoint files (*.pptx)|*.pptx"
            dlg.CheckFileExists = True
            Dim __safeDialogOwner568 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner568 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner568), dlg.ShowDialog()) = DialogResult.OK Then Return dlg.FileName
        End Using
        Return Nothing
    End Function

    Private Function PickPptxSave(srcPath As String) As String
        Using dlg As New SaveFileDialog()
            dlg.Title = "Save re-templated presentation as…"
            dlg.Filter = "PowerPoint files (*.pptx)|*.pptx"
            dlg.OverwritePrompt = True
            Dim dir = System.IO.Path.GetDirectoryName(srcPath)
            Dim name = System.IO.Path.GetFileNameWithoutExtension(srcPath) & "_retemplated.pptx"
            dlg.InitialDirectory = dir
            dlg.FileName = name
            Dim __safeDialogOwner582 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner582 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner582), dlg.ShowDialog()) = DialogResult.OK Then Return dlg.FileName
        End Using
        Return Nothing
    End Function

#End Region

#Region "Template preparation"

    ''' <summary>
    ''' Removes all sample slides from a package, leaving masters/layouts/theme/fonts intact.
    ''' </summary>
    Private Sub StripSampleSlidesFromPackage(path As String)
        Using doc As PresentationDocument = PresentationDocument.Open(path, True)
            Dim pres = doc.PresentationPart
            If pres Is Nothing OrElse pres.Presentation Is Nothing Then Return

            Dim sldIdList = pres.Presentation.SlideIdList
            If sldIdList IsNot Nothing Then
                ' Collect slide parts
                Dim sids = sldIdList.Elements(Of P.SlideId)().ToList()
                For Each sid In sids
                    Dim sp = TryCast(pres.GetPartById(sid.RelationshipId), SlidePart)
                    sid.Remove()
                    If sp IsNot Nothing Then pres.DeletePart(sp)
                Next
            Else
                EnsureSlideIdList(pres)
            End If

            ' Remove per-slide comments parts that lost their slides (safety)
            ' CommentsPart is child of SlidePart so already gone.
            pres.Presentation.Save()
        End Using
    End Sub

#End Region

#Region "Signature building"

    Private Function BuildLayoutCatalog(pptxPath As String) As List(Of LayoutSignature)
        Dim result As New List(Of LayoutSignature)
        Using doc As PresentationDocument = PresentationDocument.Open(pptxPath, False)
            Dim pres = doc.PresentationPart
            For Each sm As SlideMasterPart In pres.SlideMasterParts
                For Each lp As SlideLayoutPart In sm.SlideLayoutParts
                    Dim sig = BuildLayoutSignature(sm, lp)
                    If sig IsNot Nothing Then result.Add(sig)
                Next
            Next
        End Using
        Return result
    End Function

    Private Function BuildLayoutSignature(sm As SlideMasterPart, lp As SlideLayoutPart) As LayoutSignature
        If lp Is Nothing Then Return Nothing
        Dim sig As New LayoutSignature With {
            .Name = GetLayoutName(lp),
            .Uri = If(lp.Uri IsNot Nothing, lp.Uri.ToString(), "")
        }
        Try : sig.LayoutRelId = sm.GetIdOfPart(lp) : Catch : End Try

        Dim tree = lp.SlideLayout?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return sig

        Dim xCoords As New List(Of Long)
        For Each shp In tree.Elements(Of P.Shape)()
            Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then Continue For

            If ph.Type Is Nothing Then
                If ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then sig.HasSubTitle = True
            Else
                Select Case ph.Type.Value
                    Case PlaceholderValues.Title : sig.HasTitle = True
                    Case PlaceholderValues.CenteredTitle : sig.HasCenteredTitle = True
                    Case PlaceholderValues.SubTitle : sig.HasSubTitle = True
                    Case PlaceholderValues.Body : sig.BodyCount += 1
                    Case PlaceholderValues.Picture : sig.PicturePlaceholderCount += 1
                    Case PlaceholderValues.Object : sig.BodyCount += 1
                End Select
            End If

            Dim ofs = shp.ShapeProperties?.Transform2D?.Offset
            If ofs IsNot Nothing AndAlso ofs.X IsNot Nothing Then xCoords.Add(ofs.X.Value)
        Next

        sig.IsCoverLike = (sig.HasTitle Or sig.HasCenteredTitle) AndAlso sig.HasSubTitle AndAlso sig.BodyCount = 0
        sig.IsSectionHeaderLike = (sig.HasTitle Or sig.HasCenteredTitle) AndAlso Not sig.HasSubTitle AndAlso sig.BodyCount = 0
        sig.IsBlank = Not sig.HasTitle AndAlso Not sig.HasCenteredTitle AndAlso Not sig.HasSubTitle AndAlso sig.BodyCount = 0
        sig.IsTwoContent = sig.BodyCount >= 2 OrElse (sig.BodyCount = 1 AndAlso xCoords.Distinct().Count() >= 2)

        ' Infer by name keywords as a fallback
        Dim n = (If(sig.Name, "")).ToLowerInvariant()
        If n.Contains("section") Then sig.IsSectionHeaderLike = True
        If n.Contains("title slide") OrElse n.Contains("cover") Then sig.IsCoverLike = True
        If n.Contains("two") OrElse n.Contains("comparison") Then sig.IsTwoContent = True
        If n.Contains("blank") Then sig.IsBlank = True

        Return sig
    End Function

    Private Function BuildSlideSignatures(pptxPath As String) As List(Of SlideSignature)
        Dim list As New List(Of SlideSignature)
        Using doc As PresentationDocument = PresentationDocument.Open(pptxPath, False)
            Dim pres = doc.PresentationPart
            Dim sldIdList = pres.Presentation.SlideIdList
            If sldIdList Is Nothing Then Return list

            Dim idx As Integer = 0
            For Each sid In sldIdList.Elements(Of P.SlideId)()
                Dim sp = TryCast(pres.GetPartById(sid.RelationshipId), SlidePart)
                If sp Is Nothing Then idx += 1 : Continue For
                list.Add(BuildSlideSignature(sp, sid.Id.Value, idx))
                idx += 1
            Next
        End Using
        Return list
    End Function
    Private Function BuildSlideSignature(sp As SlidePart, slideId As UInteger, index As Integer) As SlideSignature
        Dim sig As New SlideSignature With {.Index = index, .Title = GetSlideTitle(sp)}
        sig.Key = If(String.IsNullOrWhiteSpace(sig.Title),
                     $"SID-{slideId}",
                     $"{SanitizeKey(sig.Title)}-{slideId}")

        ' Capture the source layout name for LLM matching
        Try
            sig.SourceLayoutName = GetLayoutName(sp.SlideLayoutPart)
        Catch
            sig.SourceLayoutName = String.Empty
        End Try

        Dim tree = sp.Slide?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return sig

        Dim xCoords As New List(Of Long)
        Dim bodyTexts As New List(Of String)

        For Each child In tree.ChildElements
            If TypeOf child Is P.Shape Then
                Dim shp = CType(child, P.Shape)
                sig.ShapeCount += 1
                Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                If ph IsNot Nothing Then
                    If ph.Type Is Nothing Then
                        If ph.Index IsNot Nothing AndAlso ph.Index.Value = 1UI Then sig.HasSubTitle = True
                    Else
                        Select Case ph.Type.Value
                            Case PlaceholderValues.Title, PlaceholderValues.CenteredTitle : sig.HasTitle = True
                            Case PlaceholderValues.SubTitle : sig.HasSubTitle = True
                            Case PlaceholderValues.Body, PlaceholderValues.Object
                                sig.BodyCount += 1
                                ' Capture body text for content preview
                                Try
                                    If shp.TextBody IsNot Nothing Then
                                        Dim txt = ExtractTextFromTextContainer(shp.TextBody)
                                        If Not String.IsNullOrWhiteSpace(txt) Then
                                            bodyTexts.Add(txt.Trim())
                                        End If
                                    End If
                                Catch
                                End Try
                        End Select
                    End If

                    ' Also check unnamed body placeholders (idx >= 1, no type)
                    If ph.Type Is Nothing AndAlso ph.Index IsNot Nothing AndAlso
                       ph.Index.Value >= 1UI AndAlso Not sig.HasSubTitle Then
                        ' Could be body — collect its text
                        Try
                            If shp.TextBody IsNot Nothing Then
                                Dim txt = ExtractTextFromTextContainer(shp.TextBody)
                                If Not String.IsNullOrWhiteSpace(txt) AndAlso txt.Trim().Length > 20 Then
                                    bodyTexts.Add(txt.Trim())
                                    If sig.BodyCount = 0 Then sig.BodyCount += 1
                                End If
                            End If
                        Catch
                        End Try
                    End If
                Else
                    ' Non-placeholder text shape — check for large text blocks
                    Try
                        If shp.TextBody IsNot Nothing Then
                            Dim txt = ExtractTextFromTextContainer(shp.TextBody)
                            If Not String.IsNullOrWhiteSpace(txt) AndAlso txt.Trim().Length > 40 Then
                                bodyTexts.Add(txt.Trim())
                            End If
                        End If
                    Catch
                    End Try
                End If

                Dim ofs = shp.ShapeProperties?.Transform2D?.Offset
                If ofs IsNot Nothing AndAlso ofs.X IsNot Nothing Then xCoords.Add(ofs.X.Value)
            ElseIf TypeOf child Is P.Picture Then
                sig.PictureCount += 1
            ElseIf TypeOf child Is P.GraphicFrame Then
                Dim gf = CType(child, P.GraphicFrame)
                Dim uri = gf.Graphic?.GraphicData?.Uri?.Value
                If uri IsNot Nothing Then
                    If uri.Contains("table") Then sig.TableCount += 1
                    If uri.Contains("chart") Then sig.ChartCount += 1
                    If uri.Contains("diagram") Then sig.SmartArtCount += 1
                End If
            ElseIf TypeOf child Is P.GroupShape Then
                sig.ShapeCount += 1
            End If
        Next

        sig.IsCover = (index = 0) AndAlso sig.HasTitle AndAlso sig.BodyCount = 0
        sig.IsSectionHeader = sig.HasTitle AndAlso sig.BodyCount = 0 AndAlso sig.PictureCount = 0 AndAlso sig.TableCount = 0 AndAlso sig.ChartCount = 0
        sig.IsBlank = Not sig.HasTitle AndAlso sig.BodyCount = 0 AndAlso sig.ShapeCount <= 1
        sig.TwoColumnHint = sig.BodyCount >= 2 OrElse xCoords.Distinct().Count() >= 2

        ' Build content preview: longest body text, truncated to 200 chars
        If bodyTexts.Count > 0 Then
            Dim longest = bodyTexts.OrderByDescending(Function(t) t.Length).First()
            sig.ContentPreview = If(longest.Length > 200, longest.Substring(0, 200) & "…", longest)
        Else
            sig.ContentPreview = String.Empty
        End If

        Return sig
    End Function
#End Region

#Region "Layout matching (heuristic + LLM)"

    Private Function ScoreLayoutCandidates(
            slide As SlideSignature,
            catalog As List(Of LayoutSignature)) As List(Of LayoutSignature)

        ' Work on a shallow copy so we don't mutate shared catalog objects across slides
        Dim scored As New List(Of LayoutSignature)
        For Each c In catalog
            Dim copy As New LayoutSignature With {
                .LayoutRelId = c.LayoutRelId, .Name = c.Name, .Uri = c.Uri,
                .HasTitle = c.HasTitle, .HasCenteredTitle = c.HasCenteredTitle,
                .HasSubTitle = c.HasSubTitle, .BodyCount = c.BodyCount,
                .PicturePlaceholderCount = c.PicturePlaceholderCount,
                .IsCoverLike = c.IsCoverLike, .IsSectionHeaderLike = c.IsSectionHeaderLike,
                .IsBlank = c.IsBlank, .IsTwoContent = c.IsTwoContent}
            Dim score As Double = 0

            If slide.IsCover AndAlso copy.IsCoverLike Then score += 8
            If slide.IsSectionHeader AndAlso copy.IsSectionHeaderLike Then score += 6
            If slide.IsBlank AndAlso copy.IsBlank Then score += 6

            If slide.HasTitle AndAlso (copy.HasTitle Or copy.HasCenteredTitle) Then score += 2
            If slide.HasSubTitle AndAlso copy.HasSubTitle Then score += 2
            If Not slide.HasTitle AndAlso Not (copy.HasTitle Or copy.HasCenteredTitle) Then score += 1

            If slide.BodyCount > 0 AndAlso copy.BodyCount > 0 Then score += 2
            If slide.BodyCount = 0 AndAlso copy.BodyCount = 0 Then score += 1
            score -= System.Math.Abs(slide.BodyCount - copy.BodyCount) * 0.5

            If slide.TwoColumnHint AndAlso copy.IsTwoContent Then score += 3
            If slide.PictureCount > 0 AndAlso copy.PicturePlaceholderCount > 0 Then score += 1

            Dim n = (If(copy.Name, "")).ToLowerInvariant()
            If slide.ChartCount > 0 AndAlso n.Contains("chart") Then score += 1
            If slide.TableCount > 0 AndAlso n.Contains("table") Then score += 1

            copy.LastScore = score
            scored.Add(copy)
        Next

        Return scored.OrderByDescending(Function(x) x.LastScore).ToList()
    End Function

    Private Async Function RefineMappingsWithLlm(
            signatures As List(Of SlideSignature),
            catalog As List(Of LayoutSignature),
            mappings As List(Of SlideLayoutMapping),
            options As RetemplateOptions) As Task

        ' ================================================================
        ' Build a rich JSON payload that leverages the detailed placeholder
        ' information now available from the enhanced JSON export.
        ' ================================================================

        ' ---- Layouts: include placeholder role breakdown ----
        Dim layoutsPayload = catalog.Select(Function(c)
                                                ' Build per-layout placeholder role summary from the LayoutSignature
                                                Return New With {
                Key .relId = c.LayoutRelId,
                Key .name = c.Name,
                Key .hasTitle = c.HasTitle Or c.HasCenteredTitle,
                Key .hasCenteredTitle = c.HasCenteredTitle,
                Key .hasSubTitle = c.HasSubTitle,
                Key .bodyCount = c.BodyCount,
                Key .picturePlaceholderCount = c.PicturePlaceholderCount,
                Key .isCover = c.IsCoverLike,
                Key .isSection = c.IsSectionHeaderLike,
                Key .isTwoContent = c.IsTwoContent,
                Key .isBlank = c.IsBlank
            }
                                            End Function).ToList()

        ' ---- Slides: include content preview, layout name, element counts ----
        Dim slidesPayload As New List(Of Object)
        For i As Integer = 0 To signatures.Count - 1
            Dim s = signatures(i)
            Dim top3 = ScoreLayoutCandidates(s, catalog).
                       Take(3).
                       Select(Function(c) New With {
                           Key .relId = c.LayoutRelId,
                           Key .name = c.Name,
                           Key .score = c.LastScore
                       }).ToList()

            ' Build a content preview: first 120 chars of each text block
            Dim contentPreview As String = String.Empty
            If s.ContentPreview IsNot Nothing Then
                contentPreview = s.ContentPreview
            End If

            slidesPayload.Add(New With {
                Key .key = s.Key,
                Key .index = s.Index,
                Key .title = s.Title,
                Key .sourceLayoutName = s.SourceLayoutName,
                Key .contentPreview = contentPreview,
                Key .hasTitle = s.HasTitle,
                Key .hasSubTitle = s.HasSubTitle,
                Key .bodyCount = s.BodyCount,
                Key .pictureCount = s.PictureCount,
                Key .tableCount = s.TableCount,
                Key .chartCount = s.ChartCount,
                Key .smartArtCount = s.SmartArtCount,
                Key .shapeCount = s.ShapeCount,
                Key .mediaCount = s.MediaCount,
                Key .twoColumnHint = s.TwoColumnHint,
                Key .isCover = s.IsCover,
                Key .isSectionHeader = s.IsSectionHeader,
                Key .isBlank = s.IsBlank,
                Key .candidates = top3
            })
        Next

        Dim payload As New Dictionary(Of String, Object) From {
            {"layouts", layoutsPayload},
            {"slides", slidesPayload}
        }

        Dim userPayload As String = JsonSerializer.Serialize(payload,
            New JsonSerializerOptions With {.WriteIndented = False})

        Dim systemPrompt As String =
            "You are a PowerPoint layout matcher used during a presentation re-templating operation. " &
            "You receive JSON with two top-level keys:" & vbCrLf &
            vbCrLf &
            "'layouts' – the available template layouts. Each has:" & vbCrLf &
            "  relId, name, hasTitle, hasCenteredTitle, hasSubTitle, bodyCount, " &
            "picturePlaceholderCount, isCover, isSection, isTwoContent, isBlank." & vbCrLf &
            vbCrLf &
            "'slides' – the source slides. Each has:" & vbCrLf &
            "  key, index, title, sourceLayoutName (original layout name), " &
            "contentPreview (truncated text from the main body), " &
            "hasTitle, hasSubTitle, bodyCount, pictureCount, tableCount, chartCount, " &
            "smartArtCount, shapeCount, mediaCount, twoColumnHint, " &
            "isCover, isSectionHeader, isBlank, " &
            "candidates (top 3 heuristic picks with relId, name, score)." & vbCrLf &
            vbCrLf &
            "RULES:" & vbCrLf &
            "1. For each source slide, pick the BEST layout from 'layouts' (not limited to 'candidates')." & vbCrLf &
            "2. Match on STRUCTURAL compatibility first: title/subtitle/body placeholder counts must align." & vbCrLf &
            "3. Use 'sourceLayoutName' and 'name' for semantic matching (e.g. 'Title Slide' → a cover layout)." & vbCrLf &
            "4. 'contentPreview' helps distinguish section headers (short text) from content slides (long text)." & vbCrLf &
            "5. Prefer the heuristic candidate if scores are close (difference < 2)." & vbCrLf &
            "6. Slide index 0 is almost always the cover/title slide." & vbCrLf &
            "7. Slides with tables/charts need a layout with at least 1 body placeholder." & vbCrLf &
            "8. 'isBlank' source slides should map to a blank layout if available." & vbCrLf &
            "9. Two-content slides (twoColumnHint=true) strongly prefer isTwoContent layouts." & vbCrLf &
            vbCrLf &
            "Return ONLY a JSON object:" & vbCrLf &
            "{""assignments"":[{""key"":""..."",""layoutRelId"":""..."",""reason"":""...""},…]}" & vbCrLf &
            "'reason' is a brief justification (max 10 words). Do not add explanations outside the JSON."

        Dim userPrompt As String = "<DATA>" & userPayload & "</DATA>"

        Dim reply As String = Await SLib.LLM(_context, systemPrompt, userPrompt,
                                             UseSecondAPI:=options.UseSecondaryApi,
                                             Hidesplash:=True)
        If String.IsNullOrWhiteSpace(reply) Then Return

        Dim cleaned As String = CleanJsonString(reply)
        Try
            Using jdoc As JsonDocument = JsonDocument.Parse(cleaned)
                Dim arr As JsonElement
                If Not jdoc.RootElement.TryGetProperty("assignments", arr) Then Return
                For Each el In arr.EnumerateArray()
                    Dim k As JsonElement, r As JsonElement
                    If el.TryGetProperty("key", k) AndAlso el.TryGetProperty("layoutRelId", r) Then
                        Dim key = k.GetString(), rid = r.GetString()
                        Dim m = mappings.FirstOrDefault(Function(mm) mm.SourceKey = key)
                        If m IsNot Nothing Then
                            Dim layout = catalog.FirstOrDefault(
                                Function(c) String.Equals(c.LayoutRelId, rid, StringComparison.OrdinalIgnoreCase))
                            If layout IsNot Nothing Then
                                m.LayoutRelId = layout.LayoutRelId
                                m.LayoutName = layout.Name
                                m.Rationale = "llm"

                                ' Store the LLM's reason if provided
                                Dim reasonEl As JsonElement
                                If el.TryGetProperty("reason", reasonEl) Then
                                    Dim reason = reasonEl.GetString()
                                    If Not String.IsNullOrWhiteSpace(reason) Then
                                        m.Rationale = "llm: " & reason.Trim()
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            End Using
        Catch
        End Try
    End Function

#End Region

#Region "Layout review dialog"

    ''' <summary>
    ''' Styled WinForms review: user can change assigned layout per slide.
    ''' Uses the same icon, font, padding and button style as ShowCustomVariableInputForm.
    ''' Returns False if the user cancels.
    ''' </summary>
    Private Function ShowLayoutReviewDialog(
            mappings As List(Of SlideLayoutMapping),
            catalog As List(Of LayoutSignature),
            sourcePath As String) As Boolean

        Using frm As New Form()
            frm.Text = AN & " – Review Layout Assignments"
            frm.StartPosition = FormStartPosition.CenterScreen
            frm.Width = 1060
            frm.Height = 640
            frm.MinimumSize = New System.Drawing.Size(800, 420)
            frm.FormBorderStyle = FormBorderStyle.Sizable
            frm.MaximizeBox = True
            frm.MinimizeBox = False
            frm.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            frm.AutoScaleMode = AutoScaleMode.Font
            frm.AutoScaleDimensions = New SizeF(6.0F, 13.0F)
            frm.KeyPreview = True

            ' Icon — same as all custom forms
            Try
                Dim bmpIcon As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                frm.Icon = Icon.FromHandle(bmpIcon.GetHicon())
            Catch
            End Try

            ' ---- Header label ----
            Dim headerLabel As New System.Windows.Forms.Label() With {
                .Text = "Review the layout assignment for each source slide." & vbCrLf &
                        "Source: " & System.IO.Path.GetFileName(sourcePath),
                .AutoSize = True,
                .MaximumSize = New Size(1000, 0),
                .Dock = DockStyle.Top,
                .Padding = New Padding(14, 14, 14, 10),
                .Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
            }

            ' ---- Grid wrapper panel for padding on all sides ----
            Dim gridPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(14, 10, 14, 10)
            }

            ' ---- DataGridView ----
            Dim grid As New DataGridView() With {
                .Dock = DockStyle.Fill,
                .AllowUserToAddRows = False,
                .AllowUserToDeleteRows = False,
                .RowHeadersVisible = False,
                .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                .BackgroundColor = SystemColors.Window,
                .BorderStyle = BorderStyle.FixedSingle,
                .CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
                .ColumnHeadersHeight = 52,
                .EnableHeadersVisualStyles = True,
                .Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point),
                .RowTemplate = New DataGridViewRow() With {.Height = 38}
            }

            grid.DefaultCellStyle.Padding = New Padding(6, 6, 6, 6)
            grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            grid.DefaultCellStyle.WrapMode = DataGridViewTriState.False

            grid.ColumnHeadersDefaultCellStyle.Padding = New Padding(8, 8, 8, 8)
            grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            grid.ColumnHeadersDefaultCellStyle.WrapMode = DataGridViewTriState.True
            grid.ColumnHeadersDefaultCellStyle.Font = New System.Drawing.Font("Segoe UI", 9.0F, FontStyle.Bold, GraphicsUnit.Point)

            Dim colIdx As New DataGridViewTextBoxColumn() With {
                .HeaderText = "#",
                .ReadOnly = True,
                .Width = 42,
                .MinimumWidth = 36,
                .FillWeight = 5,
                .DefaultCellStyle = New DataGridViewCellStyle() With {
                    .Alignment = DataGridViewContentAlignment.MiddleCenter}
            }
            Dim colTitle As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Source Slide Title",
                .ReadOnly = True,
                .MinimumWidth = 140,
                .FillWeight = 32
            }
            Dim colLayout As New DataGridViewComboBoxColumn() With {
                .HeaderText = "Template Layout",
                .MinimumWidth = 160,
                .FillWeight = 30,
                .FlatStyle = FlatStyle.Flat
            }
            colLayout.DataSource = catalog
            colLayout.DisplayMember = "Name"
            colLayout.ValueMember = "LayoutRelId"

            Dim colRationale As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Rationale",
                .ReadOnly = True,
                .MinimumWidth = 120,
                .FillWeight = 22
            }
            Dim colConf As New DataGridViewTextBoxColumn() With {
                .HeaderText = "Conf.",
                .ReadOnly = True,
                .Width = 54,
                .MinimumWidth = 46,
                .FillWeight = 6,
                .DefaultCellStyle = New DataGridViewCellStyle() With {
                    .Alignment = DataGridViewContentAlignment.MiddleCenter}
            }

            grid.Columns.AddRange(colIdx, colTitle, colLayout, colRationale, colConf)

            For Each m In mappings
                Dim rIdx = grid.Rows.Add()
                Dim r = grid.Rows(rIdx)
                r.Cells(0).Value = (m.SourceIndex + 1)
                r.Cells(1).Value = If(String.IsNullOrWhiteSpace(m.SourceTitle), "(untitled)", m.SourceTitle)
                r.Cells(2).Value = m.LayoutRelId
                r.Cells(3).Value = If(String.IsNullOrWhiteSpace(m.Rationale), "–", m.Rationale)
                r.Cells(4).Value = m.Confidence.ToString("0.0")
            Next

            ' ---- Button panel — same style as ShowCustomVariableInputForm ----
            Dim buttonFlow As New FlowLayoutPanel() With {
                .FlowDirection = FlowDirection.RightToLeft,
                .Dock = DockStyle.Bottom,
                .AutoSize = True,
                .AutoSizeMode = AutoSizeMode.GrowAndShrink,
                .Padding = New Padding(12, 8, 12, 12)
            }

            Dim btnApply As New Button() With {
                .Text = "Apply",
                .AutoSize = True,
                .DialogResult = DialogResult.OK,
                .Padding = New Padding(12, 4, 12, 4),
                .MinimumSize = New Size(90, 30)
            }
            Dim btnCancel As New Button() With {
                .Text = "Cancel",
                .AutoSize = True,
                .DialogResult = DialogResult.Cancel,
                .Padding = New Padding(12, 4, 12, 4),
                .MinimumSize = New Size(90, 30)
            }

            ' RightToLeft flow: add Cancel first so visual order is [Apply][Cancel]
            buttonFlow.Controls.Add(btnCancel)
            buttonFlow.Controls.Add(btnApply)
            btnApply.TabIndex = 0
            btnCancel.TabIndex = 1

            frm.AcceptButton = btnApply
            frm.CancelButton = btnCancel

            ' Ctrl+Enter → Apply
            AddHandler frm.KeyDown,
                Sub(sender As Object, e As KeyEventArgs)
                    If e.KeyCode = Keys.Enter AndAlso e.Control Then
                        btnApply.PerformClick()
                        e.SuppressKeyPress = True
                        e.Handled = True
                    End If
                End Sub

            gridPanel.Controls.Add(grid)
            frm.Controls.Add(gridPanel)
            frm.Controls.Add(headerLabel)
            frm.Controls.Add(buttonFlow)

            Dim __safeDialogOwner1184 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner1184 IsNot Nothing, frm.ShowDialog(__safeDialogOwner1184), frm.ShowDialog()) <> DialogResult.OK Then Return False

            For i As Integer = 0 To mappings.Count - 1
                Dim rid = CStr(grid.Rows(i).Cells(2).Value)
                Dim layout = catalog.FirstOrDefault(Function(c) c.LayoutRelId = rid)
                If layout IsNot Nothing AndAlso layout.LayoutRelId <> mappings(i).LayoutRelId Then
                    mappings(i).LayoutRelId = layout.LayoutRelId
                    mappings(i).LayoutName = layout.Name
                    mappings(i).Rationale = "user"
                End If
            Next
        End Using

        Return True
    End Function

    ''' <summary>
    ''' Safe variant of CloneTemplateSlide that creates a minimal slide
    ''' with placeholder stubs referencing the layout.
    ''' All visual properties (position, size, font, fill, background)
    ''' are inherited natively from the layout/master.
    ''' No layout-owned shapes or resources are copied into the slide,
    ''' eliminating entirely the dangling relationship IDs that trigger repair.
    ''' </summary>
    Private Function CloneTemplateSlideSafe(
        presPart As PresentationPart,
        layoutRelId As String
    ) As SlidePart

        Dim targetLayout As SlideLayoutPart = ResolveLayout(presPart, layoutRelId)

        If targetLayout Is Nothing Then
            targetLayout = PickCoverLikeLayout(presPart)
            If targetLayout Is Nothing Then
                targetLayout = PickDefaultLayout(presPart)
            End If
        End If

        If targetLayout Is Nothing Then
            Throw New Exception("No suitable slide layout could be resolved.")
        End If

        Dim newSlidePart As SlidePart = presPart.AddNewPart(Of SlidePart)()

        ' Build minimal ShapeTree with placeholder stubs exclusively
        Dim spTree As New P.ShapeTree(
            New P.NonVisualGroupShapeProperties(
                New P.NonVisualDrawingProperties() With {.Id = 1UI, .Name = ""},
                New P.NonVisualGroupShapeDrawingProperties(),
                New P.ApplicationNonVisualDrawingProperties()),
            New P.GroupShapeProperties(
                New A.TransformGroup(
                    New A.Offset() With {.X = 0L, .Y = 0L},
                    New A.Extents() With {.Cx = 0L, .Cy = 0L},
                    New A.ChildOffset() With {.X = 0L, .Y = 0L},
                    New A.ChildExtents() With {.Cx = 0L, .Cy = 0L})))

        Dim nextId As UInteger = 2UI
        Dim layoutTree = targetLayout.SlideLayout?.CommonSlideData?.ShapeTree

        If layoutTree IsNot Nothing Then
            For Each child In layoutTree.ChildElements
                Dim phClone As P.PlaceholderShape = Nothing
                Dim origName As String = "Placeholder"

                ' Placeholders can be defined in layout as sp, pic, or graphicFrame.
                ' PowerPoint natively initializes them ALL as p:sp stubs on the slide.
                If TypeOf child Is P.Shape Then
                    Dim shp = CType(child, P.Shape)
                    Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph IsNot Nothing Then
                        phClone = CType(ph.CloneNode(True), P.PlaceholderShape)
                        origName = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, origName)
                    End If
                ElseIf TypeOf child Is P.Picture Then
                    Dim pic = CType(child, P.Picture)
                    Dim ph = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph IsNot Nothing Then
                        phClone = CType(ph.CloneNode(True), P.PlaceholderShape)
                        origName = If(pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value, "Picture Placeholder")
                    End If
                ElseIf TypeOf child Is P.GraphicFrame Then
                    Dim gf = CType(child, P.GraphicFrame)
                    Dim ph = gf.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                    If ph IsNot Nothing Then
                        phClone = CType(ph.CloneNode(True), P.PlaceholderShape)
                        origName = If(gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value, "Object Placeholder")
                    End If
                End If

                If phClone IsNot Nothing Then
                    ' Slide-level stub strictly mapping to the layout placeholder natively
                    Dim stub As New P.Shape(
                        New P.NonVisualShapeProperties(
                            New P.NonVisualDrawingProperties() With {
                                .Id = nextId,
                                .Name = origName & " " & nextId},
                            New P.NonVisualShapeDrawingProperties(
                                New A.ShapeLocks() With {.NoGrouping = True}),
                            New P.ApplicationNonVisualDrawingProperties(phClone)),
                        New P.ShapeProperties(),
                        New P.TextBody(
                            New A.BodyProperties(),
                            New A.ListStyle(),
                            New A.Paragraph(
                                New A.EndParagraphRunProperties())))

                    spTree.AppendChild(stub)
                    nextId += 1UI
                End If
            Next
        End If

        Dim newSlide As New P.Slide(New P.CommonSlideData(spTree))

        If targetLayout.SlideLayout IsNot Nothing AndAlso
           targetLayout.SlideLayout.ColorMapOverride IsNot Nothing Then
            newSlide.ColorMapOverride = CType(
                targetLayout.SlideLayout.ColorMapOverride.CloneNode(True),
                P.ColorMapOverride)
        End If

        newSlidePart.Slide = newSlide
        newSlidePart.AddPart(targetLayout)

        Return newSlidePart
    End Function

    Private Sub KeepOnlyPlaceholderShapes(sld As P.Slide)
        Dim tree = sld?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return

        For Each child In tree.ChildElements.ToList()
            If TypeOf child Is P.NonVisualGroupShapeProperties OrElse
               TypeOf child Is P.GroupShapeProperties Then
                Continue For
            End If

            Dim keep As Boolean = False

            If TypeOf child Is P.Shape Then
                Dim shp = CType(child, P.Shape)
                Dim ph = shp.NonVisualShapeProperties?.
                             ApplicationNonVisualDrawingProperties?.
                             PlaceholderShape
                keep = (ph IsNot Nothing)
            End If

            If Not keep Then
                child.Remove()
            End If
        Next
    End Sub

#End Region

#Region "Slide processing"

    ''' <summary>
    ''' Converts a source placeholder into a free-standing text box and
    ''' appends it to the destination slide, rewriting any relationship IDs.
    ''' </summary>
    Private Sub AppendPlaceholderAsTextBox(srcShape As P.Shape,
                                           srcSp As SlidePart,
                                           dstSp As SlidePart,
                                           options As RetemplateOptions,
                                           res As RetemplateResult,
                                           spIdMap As Dictionary(Of UInteger, UInteger))
        If srcShape Is Nothing Then Return

        Dim dstTree = dstSp.Slide?.CommonSlideData?.ShapeTree
        If dstTree Is Nothing Then Return

        Dim clone = CType(srcShape.CloneNode(True), P.Shape)

        ' Strip the placeholder tag so PowerPoint treats it as a regular text box
        Dim appNv = clone.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties
        If appNv IsNot Nothing AndAlso appNv.PlaceholderShape IsNot Nothing Then
            appNv.PlaceholderShape.Remove()
        End If

        ' Assign a fresh shape ID
        Dim nv = clone.NonVisualShapeProperties?.NonVisualDrawingProperties
        If nv IsNot Nothing Then
            Dim newId As UInteger =
                dstTree.Descendants(Of P.NonVisualDrawingProperties)().
                    Select(Function(n) n.Id.Value).
                    DefaultIfEmpty(1UI).
                    Max() + 1UI
            Dim oldId = srcShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
            nv.Id = newId
            Dim currentName = If(nv.Name?.Value, "")
            nv.Name = If(String.IsNullOrWhiteSpace(currentName),
                         "Imported Text " & newId.ToString(),
                         currentName & " Imported")

            If oldId IsNot Nothing Then
                spIdMap(oldId.Value) = newId
            End If
        End If

        ' Rewrite any relationship IDs (hyperlinks, embedded images)
        RewriteRelationshipIds(clone, srcSp, dstSp, options, res)

        dstTree.AppendChild(clone)
    End Sub

    Private Sub ProcessSourceSlide(
            srcPres As PresentationPart,
            outPres As PresentationPart,
            srcSp As SlidePart,
            srcSid As P.SlideId,
            map As SlideLayoutMapping,
            options As RetemplateOptions,
            res As RetemplateResult)

        ' 1) Create new slide in output, cloned from chosen layout
        Dim newSp As SlidePart = CloneTemplateSlideSafe(outPres, map.LayoutRelId)
        Dim newId As UInteger = InsertAfter(outPres, 0UI, newSp) ' 0 = append

        Dim spIdMap As New Dictionary(Of UInteger, UInteger)

        ' Reset per-slide donor tracking
        _bodyDonorShapeIds = New HashSet(Of UInteger)

        ' 2) Transfer placeholder text (title / subtitle / body) natively via Semantic Roles
        '    This also detects non-placeholder text shapes that should fill
        '    the body placeholder and records their IDs in _bodyDonorShapeIds.
        TransferPlaceholderContent(srcSp, newSp, options, res, spIdMap)

        ' 3) Clone every non-placeholder element of source into new slide
        '    (skipping donor shapes whose text was already injected into body)
        CloneNonPlaceholderShapes(srcSp, newSp, options, spIdMap, res)

        ' 4) Notes (deep copy)
        Try : CopyNotesSlide(srcSp, newSp) : Catch ex As Exception
            res.Warnings.Add("Slide " & (map.SourceIndex + 1) & " notes: " & ex.Message)
        End Try

        ' 5) Transition + timing (with spid rewrite)
        Try : CopyTransitionAndTiming(srcSp, newSp, spIdMap) : Catch ex As Exception
            res.Warnings.Add("Slide " & (map.SourceIndex + 1) & " animations: " & ex.Message)
        End Try

        ' 6) Comments
        Try : CopyComments(srcSp, newSp, srcPres, outPres) : Catch ex As Exception
            res.Warnings.Add("Slide " & (map.SourceIndex + 1) & " comments: " & ex.Message)
        End Try

        ' 7) Slide-level attributes (hidden, showMasterSp, showMasterPhAnim)
        CopySlideAttributes(srcSp, newSp)

        ' 8) Custom XML parts attached to slide
        Try : CopyCustomXmlParts(srcSp, newSp) : Catch ex As Exception
            res.Warnings.Add("Slide " & (map.SourceIndex + 1) & " custom XML: " & ex.Message)
        End Try

        ' 9) Size rescale of custom shapes if requested
        If options.ScaleToTemplateSize Then
            RescaleCustomShapes(srcPres, outPres, newSp)
        End If

        ' 10) Hidden flag from source <p:sld show>
        Try
            If srcSp.Slide IsNot Nothing AndAlso srcSp.Slide.Show IsNot Nothing Then
                newSp.Slide.Show = DirectCast(srcSp.Slide.Show.Clone(), BooleanValue)
            End If
        Catch
        End Try
    End Sub

    ' Semantic mapping structs 
    Private Enum PhRole
        Title
        Subtitle
        Body
        Other
    End Enum

    Private Class PhInfo
        Public Property Shape As P.Shape
        Public Property PhType As PlaceholderValues?
        Public Property Index As UInteger?
        Public Property Role As PhRole
        Public Property Name As String
        Public Property TextLength As Integer
        Public Property Area As Long
        Public Property SourceOrder As Integer
    End Class

    Private Class PlaceholderRef
        Public Property PhType As PlaceholderValues?
        Public Property Index As UInteger?
        Public Property Name As String
        Public Property Area As Long
    End Class

    Private Function ExtractPlaceholderRef(child As OpenXmlElement) As PlaceholderRef
        If child Is Nothing Then Return Nothing

        Dim ph As P.PlaceholderShape = Nothing
        Dim name As String = Nothing
        Dim area As Long = 0L

        If TypeOf child Is P.Shape Then
            Dim shp = CType(child, P.Shape)
            ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            name = shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value
            area = GetShapeAreaFromElement(shp.ShapeProperties?.Transform2D)
        ElseIf TypeOf child Is P.Picture Then
            Dim pic = CType(child, P.Picture)
            ph = pic.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            name = pic.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value
            area = GetShapeAreaFromElement(pic.ShapeProperties?.Transform2D)
        ElseIf TypeOf child Is P.GraphicFrame Then
            Dim gf = CType(child, P.GraphicFrame)
            ph = gf.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            name = gf.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value
            area = GetGraphicFrameArea(gf)
        End If

        If ph Is Nothing Then Return Nothing

        Return New PlaceholderRef With {
            .PhType = If(ph.Type IsNot Nothing, CType(ph.Type.Value, PlaceholderValues?), Nothing),
            .Index = If(ph.Index IsNot Nothing, CType(ph.Index.Value, UInteger?), Nothing),
            .Name = If(name, String.Empty),
            .Area = area
        }
    End Function

    ''' <summary>
    ''' Computes area from a GraphicFrame's Transform element (P.Transform, not A.Transform2D).
    ''' </summary>
    Private Shared Function GetGraphicFrameArea(gf As P.GraphicFrame) As Long
        If gf Is Nothing OrElse gf.Transform Is Nothing Then Return 0L
        Dim ext = gf.Transform.Extents
        If ext Is Nothing OrElse ext.Cx Is Nothing OrElse ext.Cy Is Nothing Then Return 0L
        Dim rawArea As Double = CDbl(ext.Cx.Value) * CDbl(ext.Cy.Value)
        If rawArea > Long.MaxValue Then Return Long.MaxValue
        Return CLng(rawArea)
    End Function

    ''' <summary>
    ''' Computes area from a Transform2D element (works for both A.Transform2D and P.Transform).
    ''' </summary>
    Private Shared Function GetShapeAreaFromElement(xfrm As A.Transform2D) As Long
        If xfrm Is Nothing Then Return 0L
        Dim ext = xfrm.Extents
        If ext Is Nothing OrElse ext.Cx Is Nothing OrElse ext.Cy Is Nothing Then Return 0L
        Dim rawArea As Double = CDbl(ext.Cx.Value) * CDbl(ext.Cy.Value)
        If rawArea > Long.MaxValue Then Return Long.MaxValue
        Return CLng(rawArea)
    End Function

    ''' <summary>
    ''' Finds the matching layout placeholder by index first, then by name.
    ''' Returns the layout's PlaceholderRef including its geometry (area),
    ''' which is critical because slide-level stubs have no Transform2D.
    ''' </summary>
    Private Function FindMatchingLayoutPlaceholder(layoutPart As SlideLayoutPart,
                                                   idx As UInteger?,
                                                   name As String) As PlaceholderRef
        If layoutPart Is Nothing Then Return Nothing

        Dim tree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree
        If tree Is Nothing Then Return Nothing

        ' 1) Match by index (most reliable)
        If idx.HasValue Then
            For Each child In tree.ChildElements
                Dim candidate = ExtractPlaceholderRef(child)
                If candidate Is Nothing Then Continue For
                If candidate.Index.HasValue AndAlso candidate.Index.Value = idx.Value Then
                    Return candidate
                End If
            Next
        End If

        ' 2) Match by name similarity
        If Not String.IsNullOrWhiteSpace(name) Then
            For Each child In tree.ChildElements
                Dim candidate = ExtractPlaceholderRef(child)
                If candidate Is Nothing Then Continue For
                If Not String.IsNullOrWhiteSpace(candidate.Name) AndAlso
                   (candidate.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                    name.IndexOf(candidate.Name, StringComparison.OrdinalIgnoreCase) >= 0) Then
                    Return candidate
                End If
            Next
        End If

        ' 3) If both idx and name failed but we have no idx, try matching
        '    the first body-like placeholder in the layout (covers <p:ph/> with no attributes)
        If Not idx.HasValue Then
            For Each child In tree.ChildElements
                Dim candidate = ExtractPlaceholderRef(child)
                If candidate Is Nothing Then Continue For
                If candidate.PhType.HasValue AndAlso
                   (candidate.PhType.Value = PlaceholderValues.Body OrElse
                    candidate.PhType.Value = PlaceholderValues.Object) Then
                    Return candidate
                End If
            Next
        End If

        Return Nothing
    End Function

    Private Function ResolvePlaceholderRole(phType As PlaceholderValues?,
                                            idx As UInteger?,
                                            name As String) As PhRole
        If phType.HasValue Then
            Select Case phType.Value
                Case PlaceholderValues.Footer,
                     PlaceholderValues.DateAndTime,
                     PlaceholderValues.SlideNumber,
                     PlaceholderValues.Header
                    Return PhRole.Other

                Case PlaceholderValues.Title,
                     PlaceholderValues.CenteredTitle
                    Return PhRole.Title

                Case PlaceholderValues.SubTitle
                    Return PhRole.Subtitle

                Case PlaceholderValues.Body,
                     PlaceholderValues.Object
                    Return PhRole.Body
            End Select
        End If

        If idx.HasValue Then
            If idx.Value = 0UI Then Return PhRole.Title

            If idx.Value = 1UI Then
                If Not String.IsNullOrWhiteSpace(name) AndAlso
                   name.IndexOf("subtitle", StringComparison.OrdinalIgnoreCase) >= 0 Then
                    Return PhRole.Subtitle
                End If
                Return PhRole.Body
            End If

            If idx.Value >= 2UI Then Return PhRole.Body
        End If

        ' Name-based inference for shapes with no type and no index
        If Not String.IsNullOrWhiteSpace(name) Then
            Dim lName = name.ToLowerInvariant()
            If lName.Contains("subtitle") Then Return PhRole.Subtitle
            If lName.Contains("title") Then Return PhRole.Title
            If lName.Contains("content") OrElse lName.Contains("body") OrElse
               lName.Contains("text") OrElse lName.Contains("object") Then
                Return PhRole.Body
            End If
        End If

        ' BUG FIX: Do NOT return Other here unconditionally.
        ' A <p:ph/> with no attributes at all is still a placeholder.
        ' Return Body as the safest default — it's better to route unknown
        ' placeholder text into the body than to silently drop it.
        Return PhRole.Body
    End Function

    Private Function GetShapeTextLength(shp As P.Shape) As Integer
        If shp Is Nothing OrElse shp.TextBody Is Nothing Then Return 0
        Dim txt = ExtractTextFromTextContainer(shp.TextBody)
        If String.IsNullOrWhiteSpace(txt) Then Return 0
        Return txt.Trim().Length
    End Function

    ''' <summary>
    ''' Gets the area of a shape. If the shape itself has no Transform2D
    ''' (typical for slide-level placeholder stubs created by CloneTemplateSlideSafe),
    ''' falls back to reading the geometry from the matching layout placeholder.
    ''' </summary>
    Private Function GetShapeArea(shp As P.Shape,
                                  Optional layoutPart As SlideLayoutPart = Nothing) As Long
        ' 1) Try the shape's own geometry first
        Dim ext = shp?.ShapeProperties?.Transform2D?.Extents
        If ext IsNot Nothing AndAlso ext.Cx IsNot Nothing AndAlso ext.Cy IsNot Nothing Then
            Dim rawArea As Double = CDbl(ext.Cx.Value) * CDbl(ext.Cy.Value)
            If rawArea > Long.MaxValue Then Return Long.MaxValue
            Return CLng(rawArea)
        End If

        ' 2) Fall back to the matching layout placeholder's geometry
        '    This is critical: CloneTemplateSlideSafe creates stubs with
        '    empty ShapeProperties (no Transform2D), so without this
        '    fallback ALL destination body placeholders have area = 0
        '    and the "sort by largest area" logic is completely broken.
        If layoutPart IsNot Nothing AndAlso shp IsNot Nothing Then
            Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph IsNot Nothing Then
                Dim idx As UInteger? = If(ph.Index IsNot Nothing, CType(ph.Index.Value, UInteger?), Nothing)
                Dim name = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, String.Empty)
                Dim layoutRef = FindMatchingLayoutPlaceholder(layoutPart, idx, name)
                If layoutRef IsNot Nothing AndAlso layoutRef.Area > 0L Then
                    Return layoutRef.Area
                End If
            End If
        End If

        Return 0L
    End Function

    Private Function CollectContentPlaceholders(tree As P.ShapeTree,
                                                Optional layoutPart As SlideLayoutPart = Nothing) As List(Of PhInfo)
        Dim result As New List(Of PhInfo)
        If tree Is Nothing Then Return result

        Dim order As Integer = 0

        For Each shp In tree.Elements(Of P.Shape)()
            Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph Is Nothing Then
                order += 1
                Continue For
            End If

            Dim name = If(shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value, String.Empty)
            Dim phType As PlaceholderValues? =
                If(ph.Type IsNot Nothing, CType(ph.Type.Value, PlaceholderValues?), Nothing)
            Dim idx As UInteger? =
                If(ph.Index IsNot Nothing, CType(ph.Index.Value, UInteger?), Nothing)

            ' Resolve missing type/name from the layout
            Dim layoutRef = FindMatchingLayoutPlaceholder(layoutPart, idx, name)
            If Not phType.HasValue AndAlso layoutRef IsNot Nothing AndAlso layoutRef.PhType.HasValue Then
                phType = layoutRef.PhType
            End If
            If String.IsNullOrWhiteSpace(name) AndAlso layoutRef IsNot Nothing Then
                name = layoutRef.Name
            End If

            Dim role = ResolvePlaceholderRole(phType, idx, name)

            ' Skip metadata placeholders (footer/date/slidenumber)
            If role = PhRole.Other Then
                ' Only truly "other" if we have explicit metadata type
                If phType.HasValue Then
                    Select Case phType.Value
                        Case PlaceholderValues.Footer,
                             PlaceholderValues.DateAndTime,
                             PlaceholderValues.SlideNumber,
                             PlaceholderValues.Header
                            order += 1
                            Continue For
                    End Select
                End If
            End If

            Dim info As New PhInfo With {
                .Shape = shp,
                .PhType = phType,
                .Index = idx,
                .Name = name,
                .Role = role,
                .TextLength = GetShapeTextLength(shp),
                .Area = GetShapeArea(shp, layoutPart),
                .SourceOrder = order
            }

            result.Add(info)
            order += 1
        Next

        Return result
    End Function

    Private Function HasTextContent(shp As P.Shape) As Boolean
        Return GetShapeTextLength(shp) > 0
    End Function

    ''' <summary>
    ''' Finds the non-placeholder shape in the source that carries the most
    ''' text, to be injected into the destination body placeholder when
    ''' no source body placeholder was matched.
    ''' Returns the shape and its text length, or Nothing.
    ''' </summary>
    Private Function FindLargestNonPlaceholderTextShape(
            srcTree As P.ShapeTree) As P.Shape

        If srcTree Is Nothing Then Return Nothing

        Dim bestShape As P.Shape = Nothing
        Dim bestLen As Integer = 0

        For Each shp In srcTree.Elements(Of P.Shape)()
            ' Skip shapes that ARE placeholders — those are handled separately
            Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
            If ph IsNot Nothing Then Continue For

            Dim len = GetShapeTextLength(shp)
            If len > bestLen Then
                bestLen = len
                bestShape = shp
            End If
        Next

        ' Only return if there's meaningful text (not just a label)
        If bestLen >= 20 Then Return bestShape
        Return Nothing
    End Function

    ''' <summary>
    ''' Checks whether a destination body placeholder is still empty
    ''' (contains only the default empty paragraph from CloneTemplateSlideSafe).
    ''' </summary>
    Private Shared Function IsPlaceholderEmpty(shp As P.Shape) As Boolean
        If shp Is Nothing OrElse shp.TextBody Is Nothing Then Return True

        For Each para In shp.TextBody.Elements(Of A.Paragraph)()
            For Each run In para.Elements(Of A.Run)()
                If run.Text IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(run.Text.Text) Then
                    Return False
                End If
            Next
        Next

        Return True
    End Function

    Private Sub TransferPlaceholderContent(srcSp As SlidePart, dstSp As SlidePart,
                                           options As RetemplateOptions,
                                           res As RetemplateResult,
                                           spIdMap As Dictionary(Of UInteger, UInteger))
        Dim srcTree = srcSp.Slide?.CommonSlideData?.ShapeTree
        Dim dstTree = dstSp.Slide?.CommonSlideData?.ShapeTree
        If srcTree Is Nothing OrElse dstTree Is Nothing Then Return

        Dim srcPhs = CollectContentPlaceholders(srcTree, srcSp.SlideLayoutPart)
        Dim dstPhs = CollectContentPlaceholders(dstTree, dstSp.SlideLayoutPart)

        Dim usedDst As New HashSet(Of P.Shape)
        Dim matchedSrc As New HashSet(Of P.Shape)

        Dim MatchShapes =
            Sub(src As PhInfo, dst As PhInfo)
                ReplacePlaceholderText(src.Shape, dst.Shape, options)
                usedDst.Add(dst.Shape)
                matchedSrc.Add(src.Shape)

                Dim sId = src.Shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                Dim dId = dst.Shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                If sId IsNot Nothing AndAlso dId IsNot Nothing Then
                    spIdMap(sId.Value) = dId.Value
                End If
            End Sub

        ' --- 1. Title → Title ---
        Dim srcTitles = srcPhs.
            Where(Function(p) p.Role = PhRole.Title AndAlso p.TextLength > 0).
            OrderBy(Function(p) p.SourceOrder).
            ToList()

        Dim dstTitles = dstPhs.
            Where(Function(p) p.Role = PhRole.Title).
            OrderBy(Function(p) If(p.Index.HasValue, CInt(p.Index.Value), Integer.MaxValue)).
            ThenBy(Function(p) p.SourceOrder).
            ToList()

        For i As Integer = 0 To System.Math.Min(srcTitles.Count, dstTitles.Count) - 1
            MatchShapes(srcTitles(i), dstTitles(i))
        Next

        ' --- 2. Subtitle → Subtitle ---
        Dim srcSubs = srcPhs.
            Where(Function(p) p.Role = PhRole.Subtitle AndAlso p.TextLength > 0).
            OrderBy(Function(p) p.SourceOrder).
            ToList()

        Dim dstSubs = dstPhs.
            Where(Function(p) p.Role = PhRole.Subtitle AndAlso Not usedDst.Contains(p.Shape)).
            OrderBy(Function(p) If(p.Index.HasValue, CInt(p.Index.Value), Integer.MaxValue)).
            ThenBy(Function(p) p.SourceOrder).
            ToList()

        For i As Integer = 0 To System.Math.Min(srcSubs.Count, dstSubs.Count) - 1
            MatchShapes(srcSubs(i), dstSubs(i))
        Next

        ' --- 3. Body → Body (largest source text → largest dest area) ---
        Dim srcBodies = srcPhs.
            Where(Function(p) p.Role = PhRole.Body AndAlso
                              Not matchedSrc.Contains(p.Shape) AndAlso
                              p.TextLength > 0).
            OrderByDescending(Function(p) p.TextLength).
            ThenByDescending(Function(p) p.Area).
            ThenBy(Function(p) p.SourceOrder).
            ToList()

        Dim dstBodies = dstPhs.
            Where(Function(p) p.Role = PhRole.Body AndAlso Not usedDst.Contains(p.Shape)).
            OrderByDescending(Function(p) p.Area).
            ThenBy(Function(p) If(p.Index.HasValue, CInt(p.Index.Value), Integer.MaxValue)).
            ThenBy(Function(p) p.SourceOrder).
            ToList()

        ' Include empty source body placeholders if no text bodies found
        If srcBodies.Count = 0 Then
            srcBodies = srcPhs.
                Where(Function(p) p.Role = PhRole.Body AndAlso Not matchedSrc.Contains(p.Shape)).
                OrderByDescending(Function(p) p.Area).
                ThenBy(Function(p) p.SourceOrder).
                ToList()
        End If

        ' Steal subtitle or other if no dest body
        If dstBodies.Count = 0 AndAlso srcBodies.Count > 0 Then
            dstBodies = dstPhs.
                Where(Function(p) (p.Role = PhRole.Subtitle OrElse p.Role = PhRole.Other) AndAlso
                                  Not usedDst.Contains(p.Shape)).
                OrderByDescending(Function(p) p.Area).
                ThenBy(Function(p) p.SourceOrder).
                ToList()
        End If

        For i As Integer = 0 To System.Math.Min(srcBodies.Count, dstBodies.Count) - 1
            MatchShapes(srcBodies(i), dstBodies(i))
        Next

        ' --- 4. Remainder ---
        Dim remainingSrc = srcPhs.
            Where(Function(s) Not matchedSrc.Contains(s.Shape) AndAlso s.TextLength > 0).
            OrderByDescending(Function(s) s.TextLength).
            ThenByDescending(Function(s) s.Area).
            ThenBy(Function(s) s.SourceOrder).
            ToList()

        Dim remainingDst = dstPhs.
            Where(Function(d) Not usedDst.Contains(d.Shape)).
            OrderByDescending(Function(d) d.Area).
            ThenBy(Function(d) d.SourceOrder).
            ToList()

        For i As Integer = 0 To remainingSrc.Count - 1
            If i < remainingDst.Count Then
                MatchShapes(remainingSrc(i), remainingDst(i))
            Else
                AppendPlaceholderAsTextBox(remainingSrc(i).Shape, srcSp, dstSp, options, res, spIdMap)
            End If
        Next

        ' =====================================================================
        ' BUG FIX #2: If the destination still has an empty body placeholder
        ' and the source had NO body placeholder at all, find the largest
        ' non-placeholder text shape in the source and inject its content
        ' into the body placeholder. This covers the very common case where
        ' the original presentation uses freestanding text boxes instead of
        ' body placeholders.
        ' =====================================================================
        Dim emptyDstBody = dstPhs.
            Where(Function(d) d.Role = PhRole.Body AndAlso
                              Not usedDst.Contains(d.Shape) AndAlso
                              IsPlaceholderEmpty(d.Shape)).
            OrderByDescending(Function(d) d.Area).
            FirstOrDefault()

        If emptyDstBody IsNot Nothing Then
            Dim donorShape = FindLargestNonPlaceholderTextShape(srcTree)
            If donorShape IsNot Nothing Then
                ReplacePlaceholderText(donorShape, emptyDstBody.Shape, options)
                usedDst.Add(emptyDstBody.Shape)

                ' Mark this shape so CloneNonPlaceholderShapes skips it
                ' (otherwise it would appear twice: once in the placeholder
                '  and once as a cloned freestanding shape).
                Dim donorId = donorShape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                Dim dstId = emptyDstBody.Shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                If donorId IsNot Nothing AndAlso dstId IsNot Nothing Then
                    spIdMap(donorId.Value) = dstId.Value
                End If

                ' Tag the donor's shape ID in spIdMap with a special sentinel
                ' so CloneNonPlaceholderShapes knows to skip it.
                ' We use the existing spIdMap: if the source shape's ID is
                ' already in spIdMap, CloneNonPlaceholderShapes won't skip it
                ' because it only skips shapes with a <p:ph> tag.
                ' We need a different mechanism — store the donor shape reference.
                _bodyDonorShapeIds.Add(If(donorId, 0UI))
            End If
        End If
    End Sub

    ''' <summary>
    ''' Shape IDs of non-placeholder shapes whose text was injected into
    ''' the destination body placeholder. CloneNonPlaceholderShapes must
    ''' skip these to avoid duplicate content.
    ''' Reset per-slide in ProcessSourceSlide.
    ''' </summary>
    Private _bodyDonorShapeIds As New HashSet(Of UInteger)

    ''' <summary>
    ''' Copies the TextBody of a source placeholder into the destination placeholder.
    ''' Preserves paragraph levels, bullets, runs, hyperlinks, bold/italic/underline.
    ''' Strips font-size/color/typeface unless PreserveSourceTypography is enabled.
    ''' </summary>
    Private Sub ReplacePlaceholderText(src As P.Shape, dst As P.Shape, options As RetemplateOptions)
        If src.TextBody Is Nothing Then Return
        Dim clone = CType(src.TextBody.CloneNode(True), P.TextBody)

        If Not options.PreserveSourceTypography Then
            For Each rPr In clone.Descendants(Of A.RunProperties)().ToList()
                rPr.FontSize = Nothing
                For Each sf In rPr.Elements(Of A.SolidFill)().ToList()
                    sf.Remove()
                Next
                For Each lf In rPr.Elements(Of A.LatinFont)().ToList()
                    lf.Remove()
                Next
                For Each ea In rPr.Elements(Of A.EastAsianFont)().ToList()
                    ea.Remove()
                Next
                For Each cs In rPr.Elements(Of A.ComplexScriptFont)().ToList()
                    cs.Remove()
                Next
            Next
        End If

        dst.TextBody = clone
    End Sub

#End Region

#Region "Custom shape deep-copy"

    ''' <summary>
    ''' Clones every non-placeholder element from source slide into destination slide,
    ''' deep-copying all related parts (images, media, charts, SmartArt, OLE, …) and
    ''' rewriting relationship IDs. Tracks a shape-id remap for later animation rewiring.
    ''' </summary>
    Private Sub CloneNonPlaceholderShapes(
            srcSp As SlidePart,
            dstSp As SlidePart,
            options As RetemplateOptions,
            spIdMap As Dictionary(Of UInteger, UInteger),
            res As RetemplateResult)

        Dim srcTree = srcSp.Slide?.CommonSlideData?.ShapeTree
        If srcTree Is Nothing Then Return
        Dim dstTree = dstSp.Slide.CommonSlideData.ShapeTree

        Dim nextId As UInteger =
            dstTree.Descendants(Of P.NonVisualDrawingProperties)().
                Select(Function(n) n.Id.Value).DefaultIfEmpty(1UI).Max() + 1UI

        For Each child In srcTree.ChildElements.ToList()
            ' Skip placeholder shapes (already handled by TransferPlaceholderContent)
            If TypeOf child Is P.Shape Then
                Dim shp = CType(child, P.Shape)
                Dim ph = shp.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape
                If ph IsNot Nothing Then Continue For

                ' Skip non-placeholder shapes whose text was injected into
                ' the destination body placeholder (donor shapes).
                Dim shpId = shp.NonVisualShapeProperties?.NonVisualDrawingProperties?.Id?.Value
                If shpId IsNot Nothing AndAlso _bodyDonorShapeIds.Contains(shpId.Value) Then
                    Continue For
                End If
            ElseIf TypeOf child Is P.NonVisualGroupShapeProperties OrElse
                   TypeOf child Is P.GroupShapeProperties Then
                Continue For
            End If

            Try
                Dim clone As OpenXmlElement = child.CloneNode(True)
                RemapShapeIds(clone, nextId, spIdMap, srcSp)
                RewriteRelationshipIds(clone, srcSp, dstSp, options, res)
                dstTree.AppendChild(clone)
            Catch ex As Exception
                res.Warnings.Add("Element clone failed (" & child.LocalName & "): " & ex.Message)
                If options.DeepCopyChartSmartArt Then
                    Try : RasterFallback(srcSp, dstSp, child, res) : Catch ex2 As Exception
                        res.Warnings.Add("Raster fallback failed: " & ex2.Message)
                    End Try
                End If
            End Try
        Next
    End Sub

    ''' <summary>
    ''' Re-assigns unique p:nvSpPr/p:nvPicPr/p:nvGrpSpPr/p:cNvPr IDs on a cloned subtree
    ''' and records the mapping so animation spTgt references can be rewritten later.
    ''' </summary>
    Private Sub RemapShapeIds(
            clone As OpenXmlElement,
            ByRef nextId As UInteger,
            spIdMap As Dictionary(Of UInteger, UInteger),
            srcSp As SlidePart)

        For Each cNvPr In clone.Descendants(Of P.NonVisualDrawingProperties)().ToList()
            If cNvPr.Id IsNot Nothing Then
                Dim oldId = cNvPr.Id.Value
                cNvPr.Id = nextId
                If Not spIdMap.ContainsKey(oldId) Then spIdMap(oldId) = nextId
                nextId += 1UI
            End If
        Next
        ' Also the enclosing element itself if it exposes a cNvPr
        Dim selfCNv = TryGet(Of P.NonVisualDrawingProperties)(clone)
        If selfCNv IsNot Nothing AndAlso selfCNv.Id IsNot Nothing Then
            ' already handled by Descendants (inclusive of self? — Descendants excludes self; add):
            Dim oldId = selfCNv.Id.Value
            If Not spIdMap.ContainsKey(oldId) Then
                selfCNv.Id = nextId
                spIdMap(oldId) = nextId
                nextId += 1UI
            End If
        End If
    End Sub

    Private Function TryGet(Of T As OpenXmlElement)(el As OpenXmlElement) As T
        If TypeOf el Is T Then Return CType(el, T)
        Return el.GetFirstChild(Of T)()
    End Function

    ''' <summary>
    ''' Finds all attributes/elements that carry an rId (r:embed, r:link, r:id) and
    ''' replaces them with freshly minted rIds that point to deep-copied parts.
    ''' </summary>
    Private Sub RewriteRelationshipIds(
            clone As OpenXmlElement,
            srcSp As SlidePart,
            dstSp As SlidePart,
            options As RetemplateOptions,
            res As RetemplateResult)

        Const rNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        ' Walk every element; inspect every attribute in r: namespace
        Dim allNodes As New List(Of OpenXmlElement)
        allNodes.Add(clone)
        allNodes.AddRange(clone.Descendants().OfType(Of OpenXmlElement)())
        For Each node In allNodes.ToList()
            For Each attr In node.GetAttributes().ToList()
                If attr.NamespaceUri <> rNs Then Continue For
                Dim oldRid As String = attr.Value
                If String.IsNullOrEmpty(oldRid) Then Continue For

                Dim newRid As String = Nothing
                Try
                    newRid = DeepCopyRelatedPart(srcSp, dstSp, oldRid, options, res)
                Catch ex As Exception
                    res.Warnings.Add("Relationship " & oldRid & " could not be copied: " & ex.Message)
                End Try

                If Not String.IsNullOrEmpty(newRid) Then
                    node.SetAttribute(New OpenXmlAttribute(attr.Prefix, attr.LocalName, rNs, newRid))
                Else
                    ' Remove the stale attribute so the output doesn't
                    ' contain a dangling relationship reference that
                    ' fails OpenXML validation.
                    node.RemoveAttribute(attr.LocalName, rNs)
                End If
            Next
        Next
    End Sub

    ''' <summary>
    ''' Copies the part referenced by <paramref name="oldRid"/> from <paramref name="srcSp"/>
    ''' (or external hyperlink) into <paramref name="dstSp"/> deeply (including its own
    ''' child parts). Returns the new relationship id.
    ''' </summary>
    ''' <summary>
    ''' Copies the part referenced by <paramref name="oldRid"/> from <paramref name="srcSp"/>
    ''' (or external hyperlink) into <paramref name="dstSp"/> deeply (including its own
    ''' child parts). Returns the new relationship id.
    ''' </summary>
    Private Function DeepCopyRelatedPart(
            srcSp As SlidePart,
            dstSp As SlidePart,
            oldRid As String,
            options As RetemplateOptions,
            res As RetemplateResult) As String

        ' External relationships (hyperlinks, linked media)
        Dim extRel = srcSp.ExternalRelationships.FirstOrDefault(Function(r) r.Id = oldRid)
        If extRel IsNot Nothing Then
            Return dstSp.AddExternalRelationship(extRel.RelationshipType, extRel.Uri).Id
        End If
        Dim hlink = srcSp.HyperlinkRelationships.FirstOrDefault(Function(r) r.Id = oldRid)
        If hlink IsNot Nothing Then
            Return dstSp.AddHyperlinkRelationship(hlink.Uri, hlink.IsExternal).Id
        End If

        ' Internal part
        Dim srcPart As OpenXmlPart = Nothing
        Try : srcPart = srcSp.GetPartById(oldRid) : Catch : End Try
        If srcPart Is Nothing Then Return Nothing

        ' Manual stream-copy approach to avoid cross-package "closed stream" errors.
        ' Buffer the source stream into memory first, then write to the new part.
        Try
            Return StreamCopyPart(srcPart, srcSp, dstSp, options, res)
        Catch ex As Exception
            res.Warnings.Add("Part copy failed for " & oldRid & " (" & srcPart.GetType().Name & "): " & ex.Message)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Copies a single OpenXmlPart by buffering its stream into memory, creating a
    ''' matching part in the destination, and recursively copying child parts.
    ''' </summary>
    Private Function StreamCopyPart(
            srcPart As OpenXmlPart,
            srcContainer As OpenXmlPartContainer,
            dstContainer As OpenXmlPartContainer,
            options As RetemplateOptions,
            res As RetemplateResult) As String

        Dim buffer As Byte()
        Using ms As New MemoryStream()
            Using srcStream = srcPart.GetStream(FileMode.Open, FileAccess.Read)
                srcStream.CopyTo(ms)
            End Using
            buffer = ms.ToArray()
        End Using

        Dim newPart As OpenXmlPart = Nothing
        Dim contentType As String = srcPart.ContentType

        If TypeOf srcPart Is ImagePart Then
            newPart = If(TypeOf dstContainer Is SlidePart,
                         CType(dstContainer, SlidePart).AddImagePart(contentType),
                         CType(dstContainer, OpenXmlPartContainer).AddNewPart(Of ImagePart)(contentType))
        ElseIf TypeOf srcPart Is EmbeddedPackagePart Then
            newPart = dstContainer.AddNewPart(Of EmbeddedPackagePart)(contentType)
        ElseIf TypeOf srcPart Is EmbeddedObjectPart Then
            newPart = dstContainer.AddNewPart(Of EmbeddedObjectPart)(contentType)
        Else
            newPart = CreateMatchingPart(srcPart, dstContainer)
        End If

        If newPart Is Nothing Then
            Dim typeName = srcPart.GetType().Name
            If Not typeName.StartsWith("CustomXml", StringComparison.OrdinalIgnoreCase) Then
                res.Warnings.Add("Unknown part type skipped: " & typeName)
            End If
            Return Nothing
        End If

        Using dataMs As New MemoryStream(buffer, writable:=False)
            newPart.FeedData(dataMs)
        End Using

        Dim ridMap As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

        For Each childRel In srcPart.Parts.ToList()
            Try
                Dim newChildRid = StreamCopyPart(childRel.OpenXmlPart, srcPart, newPart, options, res)
                If Not String.IsNullOrWhiteSpace(newChildRid) Then
                    ridMap(childRel.RelationshipId) = newChildRid
                End If
            Catch ex As Exception
                res.Warnings.Add("Child part copy skipped (" &
                                 childRel.OpenXmlPart.GetType().Name &
                                 "): " & ex.Message)
            End Try
        Next

        For Each extR In srcPart.ExternalRelationships
            Try
                Dim newRid = newPart.AddExternalRelationship(extR.RelationshipType, extR.Uri).Id
                ridMap(extR.Id) = newRid
            Catch
            End Try
        Next

        For Each hlR In srcPart.HyperlinkRelationships
            Try
                Dim newRid = newPart.AddHyperlinkRelationship(hlR.Uri, hlR.IsExternal).Id
                ridMap(hlR.Id) = newRid
            Catch
            End Try
        Next

        RewriteRelationshipIdsInPart(newPart, ridMap)

        Try
            Return dstContainer.GetIdOfPart(newPart)
        Catch
            Return Nothing
        End Try
    End Function

    Private Sub RewriteRelationshipIdsInPart(part As OpenXmlPart,
                                             ridMap As IDictionary(Of String, String))
        If part Is Nothing OrElse ridMap Is Nothing OrElse ridMap.Count = 0 Then Return
        If part.RootElement Is Nothing Then Return

        Const rNs As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

        Dim nodes As New List(Of OpenXmlElement)
        nodes.Add(part.RootElement)
        nodes.AddRange(part.RootElement.Descendants().OfType(Of OpenXmlElement)())

        For Each node In nodes
            For Each attr In node.GetAttributes().ToList()
                If attr.NamespaceUri <> rNs Then Continue For
                If String.IsNullOrWhiteSpace(attr.Value) Then Continue For
                If Not ridMap.ContainsKey(attr.Value) Then Continue For

                node.SetAttribute(New OpenXmlAttribute(
                    attr.Prefix,
                    attr.LocalName,
                    attr.NamespaceUri,
                    ridMap(attr.Value)))
            Next
        Next
    End Sub

    ''' <summary>
    ''' Creates a new part in <paramref name="dst"/> that matches the type of <paramref name="src"/>.
    ''' Covers all common PPTX part types.
    ''' </summary>
    Private Function CreateMatchingPart(src As OpenXmlPart, dst As OpenXmlPartContainer) As OpenXmlPart
        ' Image parts
        If TypeOf src Is ImagePart Then
            Return dst.AddNewPart(Of ImagePart)(src.ContentType)
        End If

        ' Chart parts
        If TypeOf src Is ChartPart Then Return dst.AddNewPart(Of ChartPart)()

        ' Diagram / SmartArt parts
        If TypeOf src Is DiagramDataPart Then Return dst.AddNewPart(Of DiagramDataPart)()
        If TypeOf src Is DiagramLayoutDefinitionPart Then Return dst.AddNewPart(Of DiagramLayoutDefinitionPart)()
        If TypeOf src Is DiagramStylePart Then Return dst.AddNewPart(Of DiagramStylePart)()
        If TypeOf src Is DiagramColorsPart Then Return dst.AddNewPart(Of DiagramColorsPart)()

        ' Embedded OLE / packages
        If TypeOf src Is EmbeddedPackagePart Then Return dst.AddNewPart(Of EmbeddedPackagePart)(src.ContentType)
        If TypeOf src Is EmbeddedObjectPart Then Return dst.AddNewPart(Of EmbeddedObjectPart)(src.ContentType)

        ' Theme override
        If TypeOf src Is ThemeOverridePart Then Return dst.AddNewPart(Of ThemeOverridePart)()

        ' Slide layout / master (should not occur at slide level, but safety)
        If TypeOf src Is SlideLayoutPart Then Return dst.AddNewPart(Of SlideLayoutPart)()

        ' VML drawing (legacy shapes in notes etc.)
        If TypeOf src Is VmlDrawingPart Then Return dst.AddNewPart(Of VmlDrawingPart)()

        ' CustomXmlPart / CustomXmlPropertiesPart: these are already
        ' bulk-copied at presentation and slide level by CopyCustomXmlParts.
        ' Creating them here via AddNewPart triggers SparseMemoryStream
        ' spill-to-disk in the VSTO AppDomain (no IsolatedStorage identity)
        ' after enough parts accumulate. Skip silently.
        If TypeOf src Is CustomXmlPart Then Return Nothing
        If src.GetType().Name = "CustomXmlPropertiesPart" Then Return Nothing

        ' Fallback: return Nothing so the caller can warn
        Return Nothing
    End Function

    ''' <summary>
    ''' PNG rasterization fallback for elements that refuse to deep-copy.
    ''' Uses NetOffice PowerPoint to export the source slide as a PNG
    ''' and inserts the image at the element's bounding box.
    ''' </summary>
    Private Sub RasterFallback(srcSp As SlidePart, dstSp As SlidePart, element As OpenXmlElement, res As RetemplateResult)
        ' Bounding box of the element, if any
        Dim xfrm = element.Descendants(Of A.Transform2D)().FirstOrDefault()
        If xfrm Is Nothing Then Return

        Dim tmpDir = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "RedInkRetmpl_" & Guid.NewGuid().ToString("N"))
        Directory.CreateDirectory(tmpDir)
        Dim pngPath As String = Nothing

        Try
            ' Export the whole source slide (not just the element) — rough but safe
            Dim ppApp As NetOffice.PowerPointApi.Application = Nothing
            Try
                ppApp = New NetOffice.PowerPointApi.Application()
                ppApp.DisplayAlerts = NetOffice.PowerPointApi.Enums.PpAlertLevel.ppAlertsNone
                Dim filePath As String = _rasterSourcePath
                If String.IsNullOrEmpty(filePath) Then Return
                Dim pres = ppApp.Presentations.Open(filePath, [readOnly]:=True)
                Try
                    ' slide index of srcSp
                    Dim pp = srcSp.OpenXmlPackage
                    Dim sidList = CType(pp, PresentationDocument).PresentationPart.Presentation.SlideIdList
                    Dim slideIdx As Integer = 1
                    Dim i As Integer = 0
                    For Each sid In sidList.Elements(Of P.SlideId)()
                        Dim sp = CType(CType(pp, PresentationDocument).PresentationPart.GetPartById(sid.RelationshipId), SlidePart)
                        If sp Is srcSp Then slideIdx = i + 1 : Exit For
                        i += 1
                    Next
                    pngPath = System.IO.Path.Combine(tmpDir, "slide.png")
                    pres.Slides(slideIdx).Export(pngPath, "PNG")
                Finally
                    pres.Close() : pres.Dispose()
                End Try
            Finally
                If ppApp IsNot Nothing Then ppApp.Quit() : ppApp.Dispose()
            End Try

            If String.IsNullOrEmpty(pngPath) OrElse Not File.Exists(pngPath) Then Return

            ' Embed PNG as picture at xfrm bounds
            Dim ip = dstSp.AddImagePart(ImagePartType.Png)
            Using fs = File.OpenRead(pngPath)
                ip.FeedData(fs)
            End Using
            Dim rid = dstSp.GetIdOfPart(ip)
            Dim newId As UInteger =
                dstSp.Slide.CommonSlideData.ShapeTree.Descendants(Of P.NonVisualDrawingProperties)().
                    Select(Function(n) n.Id.Value).DefaultIfEmpty(1UI).Max() + 1UI

            Dim pic As New P.Picture(
                New P.NonVisualPictureProperties(
                    New P.NonVisualDrawingProperties() With {.Id = newId, .Name = "Raster " & newId},
                    New P.NonVisualPictureDrawingProperties(New A.PictureLocks() With {.NoChangeAspect = True}),
                    New P.ApplicationNonVisualDrawingProperties()),
                New P.BlipFill(
                    New A.Blip() With {.Embed = rid},
                    New A.Stretch(New A.FillRectangle())),
                New P.ShapeProperties(
                    CType(xfrm.CloneNode(True), A.Transform2D),
                    New A.PresetGeometry(New A.AdjustValueList()) With {
                        .Preset = A.ShapeTypeValues.Rectangle}))
            dstSp.Slide.CommonSlideData.ShapeTree.AppendChild(pic)

        Finally
            Try : Directory.Delete(tmpDir, True) : Catch : End Try
        End Try
    End Sub

    Private _rasterSourcePath As String

#End Region

#Region "Notes / transitions / timing / comments / attributes"

    Private Sub CopyNotesSlide(srcSp As SlidePart, dstSp As SlidePart)
        If srcSp.NotesSlidePart Is Nothing Then Return
        ' Replace any default notes part
        If dstSp.NotesSlidePart IsNot Nothing Then dstSp.DeletePart(dstSp.NotesSlidePart)

        ' Deep-copy the notes XML manually instead of AddPart (avoids cross-package stream issues)
        Dim newNotes = dstSp.AddNewPart(Of NotesSlidePart)()
        newNotes.NotesSlide = CType(srcSp.NotesSlidePart.NotesSlide.CloneNode(True), P.NotesSlide)

        ' Copy any images referenced in notes
        For Each srcImg In srcSp.NotesSlidePart.ImageParts
            Dim oldId As String = srcSp.NotesSlidePart.GetIdOfPart(srcImg)
            Dim newImg = newNotes.AddImagePart(srcImg.ContentType)
            Using imgMs As New MemoryStream()
                Using s = srcImg.GetStream(FileMode.Open, FileAccess.Read)
                    s.CopyTo(imgMs)
                End Using
                imgMs.Position = 0
                newImg.FeedData(imgMs)
            End Using
            ' Rewrite rIds in notes XML
            For Each blip In newNotes.NotesSlide.Descendants(Of DocumentFormat.OpenXml.Drawing.Blip)()
                If blip.Embed IsNot Nothing AndAlso blip.Embed.Value = oldId Then
                    blip.Embed.Value = newNotes.GetIdOfPart(newImg)
                End If
            Next
        Next

        ' Link to the output presentation's notesMaster if available
        Dim outPresDoc = CType(dstSp.OpenXmlPackage, PresentationDocument)
        Dim notesMaster = outPresDoc.PresentationPart.NotesMasterPart
        If notesMaster IsNot Nothing Then
            Try : newNotes.AddPart(notesMaster) : Catch : End Try
        End If

        ' Do NOT call newNotes.NotesSlide.Save() here.
        ' AutoSave on outDoc.Dispose() persists the notes slide once.
        ' An explicit Save adds a CompressStream wrapper to WindowsBase's
        ' "exposed streams" list for this zip entry; Dispose then opens the
        ' part for write again, which closes that wrapper's DeflateStream
        ' while the wrapper is still tracked. The next FlushExposedStreams
        ' calls Flush() on the zombie -> ObjectDisposedException.
    End Sub

    Private Sub CopyTransitionAndTiming(srcSp As SlidePart, dstSp As SlidePart, spIdMap As Dictionary(Of UInteger, UInteger))
        Dim srcSlide = srcSp.Slide
        Dim dstSlide = dstSp.Slide
        If srcSlide Is Nothing OrElse dstSlide Is Nothing Then Return

        ' Transition
        Dim srcTr = srcSlide.Elements(Of P.Transition)().FirstOrDefault()
        If srcTr IsNot Nothing Then
            Dim existing = dstSlide.Elements(Of P.Transition)().FirstOrDefault()
            If existing IsNot Nothing Then existing.Remove()
            dstSlide.AppendChild(srcTr.CloneNode(True))
        End If

        ' Timing
        Dim srcTm = srcSlide.Elements(Of P.Timing)().FirstOrDefault()
        If srcTm IsNot Nothing Then
            Dim existing = dstSlide.Elements(Of P.Timing)().FirstOrDefault()
            If existing IsNot Nothing Then existing.Remove()
            Dim clone = CType(srcTm.CloneNode(True), P.Timing)
            RemapSpIdReferences(clone, spIdMap)
            dstSlide.AppendChild(clone)
        End If
    End Sub

    Private Sub RemapSpIdReferences(root As OpenXmlElement, spIdMap As Dictionary(Of UInteger, UInteger))
        If spIdMap Is Nothing OrElse spIdMap.Count = 0 Then Return
        ' Walk every element; look for attributes named "spid" in any namespace
        Dim allNodes As New List(Of OpenXmlElement)
        allNodes.Add(root)
        allNodes.AddRange(root.Descendants().OfType(Of OpenXmlElement)())
        For Each node In allNodes
            For Each attr In node.GetAttributes().ToList()
                If Not String.Equals(attr.LocalName, "spid", StringComparison.OrdinalIgnoreCase) Then Continue For
                Dim v As UInteger
                If UInteger.TryParse(attr.Value, v) AndAlso spIdMap.ContainsKey(v) Then
                    node.SetAttribute(New OpenXmlAttribute(attr.Prefix, attr.LocalName, attr.NamespaceUri, spIdMap(v).ToString()))
                End If
            Next
        Next
    End Sub

    Private Sub CopyComments(srcSp As SlidePart, dstSp As SlidePart, srcPres As PresentationPart, outPres As PresentationPart)
        Dim srcComments = srcSp.SlideCommentsPart
        If srcComments Is Nothing Then Return

        ' Ensure output CommentAuthorsPart exists and merge authors
        Dim outAuthorsPart = outPres.CommentAuthorsPart
        If outAuthorsPart Is Nothing Then
            outAuthorsPart = outPres.AddNewPart(Of CommentAuthorsPart)()
            outAuthorsPart.CommentAuthorList = New P.CommentAuthorList()
        End If

        Dim srcAuthorsPart = srcPres.CommentAuthorsPart
        If srcAuthorsPart IsNot Nothing Then
            For Each author As P.CommentAuthor In srcAuthorsPart.CommentAuthorList.Elements(Of P.CommentAuthor)()
                Dim exists = outAuthorsPart.CommentAuthorList.Elements(Of P.CommentAuthor)().
                             Any(Function(x) x.Id.Value = author.Id.Value)
                If Not exists Then
                    outAuthorsPart.CommentAuthorList.AppendChild(Of P.CommentAuthor)(
                        CType(author.CloneNode(True), P.CommentAuthor))
                End If
            Next
            ' Do NOT call outAuthorsPart.CommentAuthorList.Save() here.
            ' CopyComments runs once per slide that has comments, but they
            ' all target the same commentAuthors.xml zip entry. Each Save
            ' opens GetStream(FileMode.Create) on that entry and registers
            ' a CompressStream wrapper in WindowsBase's exposed-streams
            ' list. After N slides we have N wrappers; the next mode
            ' change disposes their inner DeflateStreams while the
            ' wrappers are still tracked. Dispose's FlushExposedStreams
            ' then throws ObjectDisposedException.
            ' AutoSave on outDoc.Dispose() persists the merged author list
            ' exactly once.
        End If

        If dstSp.SlideCommentsPart IsNot Nothing Then dstSp.DeletePart(dstSp.SlideCommentsPart)
        Dim newComments = dstSp.AddNewPart(Of SlideCommentsPart)()

        Using srcMs As New MemoryStream()
            Using s = srcComments.GetStream(FileMode.Open, FileAccess.Read)
                s.CopyTo(srcMs)
            End Using
            srcMs.Position = 0
            newComments.FeedData(srcMs)
        End Using
    End Sub

    Private Sub EnsureCommentAuthors(srcPres As PresentationPart, outPres As PresentationPart)
        If srcPres.CommentAuthorsPart IsNot Nothing AndAlso outPres.CommentAuthorsPart Is Nothing Then
            Dim p = outPres.AddNewPart(Of CommentAuthorsPart)()
            p.CommentAuthorList = New P.CommentAuthorList()
        End If
    End Sub

    Private Sub CopySlideAttributes(srcSp As SlidePart, dstSp As SlidePart)
        Dim s = srcSp.Slide, d = dstSp.Slide
        If s Is Nothing OrElse d Is Nothing Then Return
        If s.ShowMasterShapes IsNot Nothing Then d.ShowMasterShapes = DirectCast(s.ShowMasterShapes.Clone(), BooleanValue)
        If s.ShowMasterPlaceholderAnimations IsNot Nothing Then d.ShowMasterPlaceholderAnimations = DirectCast(s.ShowMasterPlaceholderAnimations.Clone(), BooleanValue)
    End Sub

#End Region

#Region "Presentation-level helpers"

    Private Sub CopyCustomXmlParts(srcContainer As OpenXmlPartContainer, dstContainer As OpenXmlPartContainer)
        For Each cxp In srcContainer.GetPartsOfType(Of CustomXmlPart)()
            Try
                Using srcMs As New MemoryStream()
                    Using s = cxp.GetStream(FileMode.Open, FileAccess.Read)
                        s.CopyTo(srcMs)
                    End Using
                    srcMs.Position = 0
                    Dim newCxp = dstContainer.AddNewPart(Of CustomXmlPart)("application/xml")
                    newCxp.FeedData(srcMs)
                End Using
            Catch
            End Try
        Next
    End Sub

    Private Sub CopyEmbeddedFonts(srcPres As PresentationPart, outPres As PresentationPart)
        ' Template fonts are already present because output IS the template.
        ' This adds any font used in source but missing from template.
        For Each srcFp In srcPres.GetPartsOfType(Of FontPart)()
            Dim already = outPres.GetPartsOfType(Of FontPart)().
                Any(Function(fp) fp.ContentType = srcFp.ContentType AndAlso
                                 fp.Uri.ToString() = srcFp.Uri.ToString())
            If Not already Then
                Try
                    Using srcMs As New MemoryStream()
                        Using s = srcFp.GetStream(FileMode.Open, FileAccess.Read)
                            s.CopyTo(srcMs)
                        End Using
                        srcMs.Position = 0
                        Dim newFp = outPres.AddNewPart(Of FontPart)(srcFp.ContentType)
                        newFp.FeedData(srcMs)
                    End Using
                Catch
                End Try
            End If
        Next
    End Sub

    Private Sub HandleSlideSize(srcPres As PresentationPart, outPres As PresentationPart, options As RetemplateOptions, res As RetemplateResult)
        Dim ss = srcPres.Presentation.SlideSize
        Dim ts = outPres.Presentation.SlideSize
        If ss Is Nothing OrElse ts Is Nothing Then Return
        If ss.Cx.Value = ts.Cx.Value AndAlso ss.Cy.Value = ts.Cy.Value Then Return

        If options.ScaleToTemplateSize Then
            res.Warnings.Add("Slide size differs; source shapes will be scaled to fit the template.")
        Else
            res.Warnings.Add(String.Format(
                "Slide size differs (source {0}x{1} EMU vs template {2}x{3} EMU). Content may overflow.",
                ss.Cx.Value, ss.Cy.Value, ts.Cx.Value, ts.Cy.Value))
        End If
    End Sub

    Private Sub RescaleCustomShapes(srcPres As PresentationPart, outPres As PresentationPart, dstSp As SlidePart)
        Dim ss = srcPres.Presentation.SlideSize
        Dim ts = outPres.Presentation.SlideSize
        If ss Is Nothing OrElse ts Is Nothing Then Return
        Dim sx As Double = CDbl(ts.Cx.Value) / ss.Cx.Value
        Dim sy As Double = CDbl(ts.Cy.Value) / ss.Cy.Value

        For Each xfrm In dstSp.Slide.Descendants(Of A.Transform2D)()
            Dim ofs = xfrm.Offset
            Dim ext = xfrm.Extents
            If ofs IsNot Nothing Then
                If ofs.X IsNot Nothing Then ofs.X.Value = CLng(ofs.X.Value * sx)
                If ofs.Y IsNot Nothing Then ofs.Y.Value = CLng(ofs.Y.Value * sy)
            End If
            If ext IsNot Nothing Then
                If ext.Cx IsNot Nothing Then ext.Cx.Value = CLng(ext.Cx.Value * sx)
                If ext.Cy IsNot Nothing Then ext.Cy.Value = CLng(ext.Cy.Value * sy)
            End If
        Next
        ' Do NOT call dstSp.Slide.Save() here. Same reason as in
        ' ProcessSourceSlide: AutoSave on outDoc.Dispose() persists the
        ' slide exactly once. An explicit Save here causes the zombie
        ' CompressStream / disposed DeflateStream pattern that throws
        ' ObjectDisposedException in FlushExposedStreams during Dispose.
    End Sub


#End Region

#Region "Manual ZIP writer (bypasses WindowsBase _exposedStreams bug)"

    ''' <summary>
    ''' Enumerates parts and relationships entirely through the OpenXML
    ''' SDK (no System.IO.Packaging.Package access), then writes a fresh
    ''' .pptx via System.IO.Compression.ZipArchive. This bypasses the
    ''' WindowsBase SaveContainer -> FlushExposedStreams path that
    ''' throws ObjectDisposedException for this codebase.
    ''' </summary>
    Private Sub WriteOpenXmlPackageAsZip(
            outDoc As OpenXmlPackage,
            tempClonePath As String,
            outputPath As String,
            res As RetemplateResult)

        ' Absolute part URI -> OpenXmlPart
        Dim allParts As New Dictionary(Of String, OpenXmlPart)(StringComparer.OrdinalIgnoreCase)
        CollectAllOpenXmlParts(outDoc, allParts)

        ' Zip entry path -> bytes
        Dim entries As New Dictionary(Of String, Byte())(StringComparer.OrdinalIgnoreCase)

        ' --- 1) Part contents ---
        For Each kvp In allParts
            Dim oxp As OpenXmlPart = kvp.Value
            Dim absUri As String = kvp.Key
            Dim zipPath As String = absUri.TrimStart("/"c)
            Dim bytes As Byte() = Nothing

            If oxp.RootElement IsNot Nothing Then
                ' In-memory DOM serialization. No package stream opened,
                ' no _exposedStreams entry created; cannot trigger the bug.
                Using ms As New MemoryStream()
                    Dim xws As New Xml.XmlWriterSettings() With {
                        .Encoding = New System.Text.UTF8Encoding(False),
                        .CloseOutput = False,
                        .OmitXmlDeclaration = False
                    }
                    Using xw As Xml.XmlWriter = Xml.XmlWriter.Create(ms, xws)
                        oxp.RootElement.WriteTo(xw)
                    End Using
                    bytes = ms.ToArray()
                End Using
            Else
                ' No loaded DOM. Try reading via the SDK's stream API.
                ' ms.ToArray() captures bytes BEFORE the Using disposes,
                ' so if the close cascade throws the zombie
                ' ObjectDisposedException we already have our data.
                Try
                    Using s = oxp.GetStream(FileMode.Open, FileAccess.Read)
                        Using ms As New MemoryStream()
                            s.CopyTo(ms)
                            bytes = ms.ToArray()
                        End Using
                    End Using
                Catch
                    ' Fall through to disk fallback below.
                End Try

                ' Fallback: pristine bytes from the temp clone on disk.
                ' Valid because AutoSave = False means the polluted
                ' package never wrote anything back to tempClonePath.
                If bytes Is Nothing Then
                    bytes = TryReadZipEntryBytes(tempClonePath, zipPath)
                End If

                If bytes Is Nothing Then
                    res.Warnings.Add("Manual-zip: could not retrieve bytes for part " & zipPath & "; entry skipped.")
                    Continue For
                End If
            End If

            entries(zipPath) = bytes
        Next

        ' --- 2) Root relationships (_rels/.rels) ---
        Dim rootRels As List(Of RelInfo) = EnumeratePackageRelationships(outDoc).ToList()
        If rootRels.Count > 0 Then
            entries("_rels/.rels") = BuildRelsXmlBytes(rootRels, String.Empty)
        End If

        ' --- 3) Per-part relationships ---
        For Each kvp In allParts
            Dim oxp As OpenXmlPart = kvp.Value
            Dim absUri As String = kvp.Key
            Dim rels As List(Of RelInfo) = EnumeratePartRelationships(oxp).ToList()
            If rels.Count = 0 Then Continue For
            entries(GetRelsEntryPathFor(absUri)) = BuildRelsXmlBytes(rels, absUri)
        Next

        ' --- 4) [Content_Types].xml ---
        entries("[Content_Types].xml") = BuildContentTypesXmlBytes(allParts.Values)

        ' --- 5) Write the zip ---
        If System.IO.File.Exists(outputPath) Then
            System.IO.File.Delete(outputPath)
        End If

        Using fs = System.IO.File.Create(outputPath)
            Using zip As New System.IO.Compression.ZipArchive(
                    fs, System.IO.Compression.ZipArchiveMode.Create)

                For Each kvp In entries
                    Dim entry = zip.CreateEntry(
                        kvp.Key,
                        System.IO.Compression.CompressionLevel.Optimal)
                    Using es = entry.Open()
                        es.Write(kvp.Value, 0, kvp.Value.Length)
                    End Using
                Next
            End Using
        End Using
    End Sub

    Private Class RelInfo
        Public Id As String
        Public Type As String
        Public TargetAbsoluteUri As String  ' "/ppt/slides/slide1.xml" for internal, or full URI for external
        Public External As Boolean
    End Class

    Private Sub CollectAllOpenXmlParts(
            pkg As OpenXmlPackage,
            map As Dictionary(Of String, OpenXmlPart))
        Dim visited As New HashSet(Of OpenXmlPart)()
        ' NOTE: loop variable intentionally named 'pair' (not 'p') to
        ' avoid colliding with the 'P' namespace alias for
        ' DocumentFormat.OpenXml.Presentation in this file.
        For Each pair As IdPartPair In pkg.Parts
            CollectOpenXmlPartRec(pair.OpenXmlPart, map, visited)
        Next
    End Sub

    Private Sub CollectOpenXmlPartRec(
            part As OpenXmlPart,
            map As Dictionary(Of String, OpenXmlPart),
            visited As HashSet(Of OpenXmlPart))
        If part Is Nothing OrElse Not visited.Add(part) Then Return
        map(part.Uri.ToString()) = part
        For Each childPair As IdPartPair In part.Parts
            CollectOpenXmlPartRec(childPair.OpenXmlPart, map, visited)
        Next
    End Sub

    Private Iterator Function EnumeratePackageRelationships(
            pkg As OpenXmlPackage) As IEnumerable(Of RelInfo)
        For Each pair As IdPartPair In pkg.Parts
            Yield New RelInfo With {
                .Id = pair.RelationshipId,
                .Type = pair.OpenXmlPart.RelationshipType,
                .TargetAbsoluteUri = pair.OpenXmlPart.Uri.ToString(),
                .External = False
            }
        Next
        For Each er In pkg.ExternalRelationships
            Yield New RelInfo With {
                .Id = er.Id,
                .Type = er.RelationshipType,
                .TargetAbsoluteUri = er.Uri.OriginalString,
                .External = True
            }
        Next
        For Each hr In pkg.HyperlinkRelationships
            Yield New RelInfo With {
                .Id = hr.Id,
                .Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                .TargetAbsoluteUri = hr.Uri.OriginalString,
                .External = hr.IsExternal
            }
        Next
    End Function

    Private Iterator Function EnumeratePartRelationships(
            part As OpenXmlPart) As IEnumerable(Of RelInfo)
        For Each pair As IdPartPair In part.Parts
            Yield New RelInfo With {
                .Id = pair.RelationshipId,
                .Type = pair.OpenXmlPart.RelationshipType,
                .TargetAbsoluteUri = pair.OpenXmlPart.Uri.ToString(),
                .External = False
            }
        Next
        For Each er In part.ExternalRelationships
            Yield New RelInfo With {
                .Id = er.Id,
                .Type = er.RelationshipType,
                .TargetAbsoluteUri = er.Uri.OriginalString,
                .External = True
            }
        Next
        For Each hr In part.HyperlinkRelationships
            Yield New RelInfo With {
                .Id = hr.Id,
                .Type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                .TargetAbsoluteUri = hr.Uri.OriginalString,
                .External = hr.IsExternal
            }
        Next
    End Function

    Private Function GetRelsEntryPathFor(partAbsoluteUri As String) As String
        Dim p = partAbsoluteUri.TrimStart("/"c)
        Dim slash = p.LastIndexOf("/"c)
        If slash < 0 Then
            Return "_rels/" & p & ".rels"
        End If
        Return p.Substring(0, slash) & "/_rels/" & p.Substring(slash + 1) & ".rels"
    End Function

    Private Function BuildRelsXmlBytes(
            rels As IEnumerable(Of RelInfo),
            sourcePartAbsoluteUri As String) As Byte()
        Dim sb As New System.Text.StringBuilder()
        sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        sb.Append("<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">")
        For Each r In rels
            Dim targetForXml As String
            If r.External Then
                targetForXml = r.TargetAbsoluteUri
            Else
                targetForXml = MakeRelativeTarget(sourcePartAbsoluteUri, r.TargetAbsoluteUri)
            End If
            sb.Append("<Relationship Id=""")
            sb.Append(System.Security.SecurityElement.Escape(r.Id))
            sb.Append(""" Type=""")
            sb.Append(System.Security.SecurityElement.Escape(r.Type))
            sb.Append(""" Target=""")
            sb.Append(System.Security.SecurityElement.Escape(targetForXml))
            sb.Append("""")
            If r.External Then
                sb.Append(" TargetMode=""External""")
            End If
            sb.Append("/>")
        Next
        sb.Append("</Relationships>")
        Return System.Text.Encoding.UTF8.GetBytes(sb.ToString())
    End Function

    ''' <summary>
    ''' Computes the .rels Target value: absolute-from-root (no leading
    ''' slash) for the package root rels, otherwise relative to the
    ''' source part's folder (using '../' when needed).
    ''' </summary>
    Private Function MakeRelativeTarget(
            sourcePartAbsoluteUri As String,
            targetPartAbsoluteUri As String) As String

        Dim tgt = targetPartAbsoluteUri.TrimStart("/"c)
        If String.IsNullOrEmpty(sourcePartAbsoluteUri) Then
            Return tgt
        End If

        ' Use Uri.MakeRelativeUri with a dummy scheme/host to compute
        ' a correct relative path (including ../ segments).
        Dim baseUri As New Uri("http://x" & sourcePartAbsoluteUri)
        Dim tgtUri As New Uri("http://x" & targetPartAbsoluteUri)
        Return baseUri.MakeRelativeUri(tgtUri).ToString()
    End Function

    Private Function BuildContentTypesXmlBytes(
            parts As IEnumerable(Of OpenXmlPart)) As Byte()

        ' Build Default Extension entries from actual parts.
        ' Group by file extension; if all parts with that extension share
        ' one content type, emit a Default (matching real PPTX behaviour).
        Dim defaults As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        defaults("rels") = "application/vnd.openxmlformats-package.relationships+xml"
        defaults("xml") = "application/xml"

        Dim overrideParts As New List(Of KeyValuePair(Of String, String))

        Dim byExt = parts.GroupBy(Function(p)
                                      Dim u = p.Uri.ToString()
                                      Dim dot = u.LastIndexOf("."c)
                                      Return If(dot >= 0, u.Substring(dot + 1).ToLowerInvariant(), "")
                                  End Function)

        For Each grp In byExt
            Dim ext = grp.Key
            If String.IsNullOrEmpty(ext) Then
                ' No extension — must use Override
                For Each oxp In grp
                    overrideParts.Add(New KeyValuePair(Of String, String)(
                        oxp.Uri.ToString(), oxp.ContentType))
                Next
                Continue For
            End If

            Dim distinctCt = grp.Select(Function(p) p.ContentType).Distinct().ToList()

            If distinctCt.Count = 1 Then
                ' All parts with this extension share one content type → Default
                If Not defaults.ContainsKey(ext) Then
                    defaults(ext) = distinctCt(0)
                ElseIf defaults(ext) <> distinctCt(0) Then
                    ' Extension already has a different default — use Override
                    For Each oxp In grp
                        overrideParts.Add(New KeyValuePair(Of String, String)(
                            oxp.Uri.ToString(), oxp.ContentType))
                    Next
                End If
                ' else: already have matching default, no Override needed
            Else
                ' Mixed content types for this extension — Override for each
                For Each oxp In grp
                    overrideParts.Add(New KeyValuePair(Of String, String)(
                        oxp.Uri.ToString(), oxp.ContentType))
                Next
            End If
        Next

        Dim sb As New System.Text.StringBuilder()
        sb.Append("<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>")
        sb.Append("<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">")

        For Each kvp In defaults
            sb.Append("<Default Extension=""")
            sb.Append(System.Security.SecurityElement.Escape(kvp.Key))
            sb.Append(""" ContentType=""")
            sb.Append(System.Security.SecurityElement.Escape(kvp.Value))
            sb.Append("""/>")
        Next

        For Each kvp In overrideParts
            sb.Append("<Override PartName=""")
            sb.Append(System.Security.SecurityElement.Escape(kvp.Key))
            sb.Append(""" ContentType=""")
            sb.Append(System.Security.SecurityElement.Escape(kvp.Value))
            sb.Append("""/>")
        Next

        sb.Append("</Types>")
        Return System.Text.Encoding.UTF8.GetBytes(sb.ToString())
    End Function

    Private Function TryReadZipEntryBytes(zipPath As String, entryName As String) As Byte()
        Try
            Using fs = System.IO.File.OpenRead(zipPath)
                Using zip As New System.IO.Compression.ZipArchive(
                        fs, System.IO.Compression.ZipArchiveMode.Read)
                    Dim entry = zip.GetEntry(entryName)
                    If entry Is Nothing Then Return Nothing
                    Using es = entry.Open()
                        Using ms As New MemoryStream()
                            es.CopyTo(ms)
                            Return ms.ToArray()
                        End Using
                    End Using
                End Using
            End Using
        Catch
            Return Nothing
        End Try
    End Function

#End Region



End Class
