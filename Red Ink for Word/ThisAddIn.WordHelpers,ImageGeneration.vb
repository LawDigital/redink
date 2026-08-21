' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WordHelpers.ImageGeneration.vb
' Purpose: Interactive image generation using a configured "ImageGeneration"
'          special task model. Prompts the user for a description, optionally
'          attaches a reference image via (file), calls the LLM, displays
'          the result in an HTML window, and offers insertion into the active
'          Word document.
'
' Architecture:
'  - Availability Check: Uses IsImageGenerationAvailable() to probe for an
'    "ImageGeneration" model in the alternate model INI and detect
'    APICall_Object support for image-editing workflows.
'  - Prompt Collection: ShowCustomInputBox (multi-line) with Ctrl-P recall
'    of the last image generation prompt (persisted in My.Settings) and an
'    optional 📎 insert button for the (file) trigger (only shown when the
'    model has object support).
'  - File Object: When (file) is present, DragDropForm (FileOnly) collects
'    the reference image path, filtered to common image formats.
'  - LLM Invocation: Switches to the ImageGeneration model via
'    GetSpecialTaskModel, calls LLM() with the prompt, and restores the
'    original config afterwards.
'  - Result Display: Converts Markdown model text to HTML via Markdig,
'    parses "Image saved to:" from the LLM response, shows the image and
'    rendered text in ShowHTMLCustomMessageBox with buttons for copying
'    the path and inserting the image into Word.
'  - Document Insertion: Adds an InlineShape at the current selection via
'    the Word object model.
'
' Dependencies:
'  - SharedLibrary (LLM, model selection, UI dialogs, ImageDecoder)
'  - DragDropForm (file selection for reference images)
'  - Markdig (Markdown to HTML conversion)
'  - Microsoft.Office.Interop.Word (document insertion)
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks
Imports Markdig
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports Word = Microsoft.Office.Interop.Word

Partial Public Class ThisAddIn

    ''' <summary>
    ''' The (file) trigger text used within the image generation prompt to indicate
    ''' that a reference image should be attached.
    ''' </summary>
    Private Const ImageGen_FileTrigger As String = "(file)"

    ''' <summary>
    ''' Entry point for interactive image generation. Checks availability of the
    ''' "ImageGeneration" special task model, collects a prompt from the user,
    ''' optionally attaches a reference image, calls the LLM, and displays the result.
    ''' </summary>
    Public Async Sub GenerateImage()
        If INILoadFail() Then Return

        Try
            ' ── 1) Check whether an ImageGeneration model is configured ──

            Dim imgModelHasObjectCall As Boolean = False
            Dim imgModelAvailable As Boolean = IsImageGenerationAvailable(_context, imgModelHasObjectCall)

            If Not imgModelAvailable Then
                ShowCustomMessageBox(
                    "No ""ImageGeneration"" special task model is configured." & vbCrLf & vbCrLf &
                    "To use this feature, add a model section with ""ImageGeneration=True"" " &
                    "to your alternate models INI file, or call up an image generation capable " &
                    "model via ""Freestyle (2nd)"".")
                Return
            End If

            ' ── 2) Collect the image description prompt ──

            Dim insertButtonsList As New System.Collections.Generic.List(Of
                System.Tuple(Of String, String, String))()

            If imgModelHasObjectCall Then
                insertButtonsList.Add(
                    System.Tuple.Create("📎", "Attach a reference image for editing/modification (file)", ImageGen_FileTrigger))
            End If

            ' Grab the current Word selection and/or entire document text for insert buttons
            Dim selectionTextForInsert As String = ""
            Dim documentTextForInsert As String = ""
            Dim hasExplicitSelection As Boolean = False

            Try
                Dim application As Word.Application = Globals.ThisAddIn.Application
                If application IsNot Nothing AndAlso application.Documents IsNot Nothing AndAlso application.Documents.Count > 0 Then
                    Dim sel As Word.Selection = application.Selection

                    ' Check for an explicit selection (not just an insertion point)
                    If sel IsNot Nothing AndAlso sel.Range IsNot Nothing AndAlso sel.Start <> sel.End Then
                        hasExplicitSelection = True

                        ' When the selection contains table cells, sel.Text only returns
                        ' the first cell. Iterate sel.Cells to capture all cell content.
                        Try
                            If sel.Cells IsNot Nothing AndAlso sel.Cells.Count > 1 Then
                                Dim sb As New System.Text.StringBuilder()
                                Dim lastRowIndex As Integer = -1
                                For Each cell As Word.Cell In sel.Cells
                                    Dim cellRange As Word.Range = cell.Range.Duplicate
                                    ' Exclude the end-of-cell marker character
                                    cellRange.End -= 1
                                    Dim cellText As String = cellRange.Text
                                    If cell.RowIndex <> lastRowIndex Then
                                        If sb.Length > 0 Then sb.AppendLine()
                                        lastRowIndex = cell.RowIndex
                                    Else
                                        sb.Append(vbTab)
                                    End If
                                    sb.Append(If(cellText, ""))
                                Next
                                selectionTextForInsert = sb.ToString()
                            Else
                                selectionTextForInsert = sel.Text
                            End If
                        Catch
                            ' sel.Cells throws COM exception for non-table selections
                            selectionTextForInsert = sel.Text
                        End Try
                    End If

                    ' Always grab the full document text
                    documentTextForInsert = application.ActiveDocument.Content.Text
                End If
            Catch
                ' Silently ignore — buttons simply won't appear
            End Try

            ' Selection button (only when text is explicitly selected)
            If hasExplicitSelection AndAlso Not String.IsNullOrWhiteSpace(selectionTextForInsert) Then
                Dim selLabel As String = If(selectionTextForInsert.Length > 200,
                    "Insert current selection (" & selectionTextForInsert.Length.ToString("N0") & " chars)",
                    "Insert current selection")
                insertButtonsList.Add(
                    System.Tuple.Create("📋", selLabel, selectionTextForInsert))
            End If

            ' Document button (always, when there is document content)
            If Not String.IsNullOrWhiteSpace(documentTextForInsert) Then
                Dim docLabel As String = If(documentTextForInsert.Length > 200,
                    "Insert entire document text (" & documentTextForInsert.Length.ToString("N0") & " chars)",
                    "Insert entire document text")
                insertButtonsList.Add(
                    System.Tuple.Create("📄", docLabel, documentTextForInsert))
            End If

            Dim insertButtons As System.Tuple(Of String, String, String)() = Nothing
            If insertButtonsList.Count > 0 Then
                insertButtons = insertButtonsList.ToArray()
            End If

            Dim lastPromptHint As String = ""
            If Not String.IsNullOrWhiteSpace(My.Settings.LastImageGenPrompt) Then
                lastPromptHint = " Ctrl-P inserts your last prompt."
            End If

            Dim promptHintParts As New System.Collections.Generic.List(Of String)()
            If hasExplicitSelection Then promptHintParts.Add("'Clipboard' button below to insert the current selection")
            promptHintParts.Add("'Document' button to insert the entire document text")

            Dim prompt As String = SLib.ShowCustomInputBox(
                "Describe the image you want to generate." &
                If(imgModelHasObjectCall,
                   " Add '" & ImageGen_FileTrigger & "' (or use the clip button) to include a reference image for editing.",
                   "") &
                " Use the " & String.Join(" or ", promptHintParts) & " button." &
                lastPromptHint,
                AN & " Image Generation",
                False,
                "",
                My.Settings.LastImageGenPrompt,
                Nothing,
                insertButtons).Trim()

            ' ESC or empty → abort
            If prompt = "ESC" OrElse String.IsNullOrWhiteSpace(prompt) Then Return

            ' Persist for Ctrl-P recall (same pattern as Commands.Freestyle)
            My.Settings.LastImageGenPrompt = prompt
            My.Settings.Save()

            ' ── 3) Handle (file) trigger ──

            Dim fileObject As String = ""
            If imgModelHasObjectCall AndAlso prompt.IndexOf(ImageGen_FileTrigger, StringComparison.OrdinalIgnoreCase) >= 0 Then
                ' Strip the trigger text from the prompt
                prompt = prompt.Replace(ImageGen_FileTrigger, "").Trim()

                ' Prompt user for a reference image file
                DragDropFormLabel = "Image files (jpg, png, gif, bmp, webp, tiff)."
                DragDropFormFilter =
                    "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.webp;*.tiff;*.tif|" &
                    "JPEG (*.jpg;*.jpeg)|*.jpg;*.jpeg|" &
                    "PNG (*.png)|*.png|" &
                    "GIF (*.gif)|*.gif|" &
                    "BMP (*.bmp)|*.bmp|" &
                    "WebP (*.webp)|*.webp|" &
                    "TIFF (*.tiff;*.tif)|*.tiff;*.tif|" &
                    "All Files (*.*)|*.*"
                fileObject = GetFileName()
                DragDropFormLabel = ""
                DragDropFormFilter = ""

                If String.IsNullOrWhiteSpace(fileObject) Then
                    ShowCustomMessageBox(
                        "No reference image was selected — aborting." & vbCrLf &
                        "You can try again (use Ctrl-P to re-insert your prompt).")
                    Return
                End If
            End If

            ' ── 4) Switch to ImageGeneration model and call LLM ──

            Dim imgModelApplied As Boolean = GetSpecialTaskModel(
                _context, INI_AlternateModelPath, "ImageGeneration")

            If Not imgModelApplied Then
                If originalConfigLoaded Then
                    RestoreDefaults(_context, originalConfig)
                    originalConfigLoaded = False
                End If
                ShowCustomMessageBox("Failed to apply the ImageGeneration model. Please check your configuration.")
                Return
            End If

            Dim llmResult As String
            Try
                llmResult = Await LLM(
                    "",
                    prompt,
                    Model:="",
                    Temperature:="",
                    Timeout:=0,
                    UseSecondAPI:=True,
                    Hidesplash:=False,
                    AddUserPrompt:="",
                    FileObject:=fileObject)
            Finally
                ' Always restore the original model config
                If originalConfigLoaded Then
                    RestoreDefaults(_context, originalConfig)
                    originalConfigLoaded = False
                End If
            End Try

            If String.IsNullOrWhiteSpace(llmResult) Then
                ShowCustomMessageBox("The image generation model did not return a result.")
                Return
            End If

            ' ── 5) Parse result and display ──

            Dim imagePath As String = ""
            Dim modelText As String = llmResult

            ' Look for "Image saved to: <path>" in the response
            Dim saveMatch As Match = Regex.Match(llmResult, "Image saved to:\s*(.+?)(?:\r?\n|$)", RegexOptions.IgnoreCase)
            If saveMatch.Success Then
                imagePath = saveMatch.Groups(1).Value.Trim()
                ' Unescape doubled backslashes that HandleObject produces
                imagePath = imagePath.Replace("\\", "\")
                ' Remove the "Image saved to:" line from the display text
                modelText = llmResult.Substring(0, saveMatch.Index).Trim()
                If saveMatch.Index + saveMatch.Length < llmResult.Length Then
                    modelText = (modelText & vbCrLf &
                        llmResult.Substring(saveMatch.Index + saveMatch.Length)).Trim()
                End If
            End If

            ' Convert Markdown model text to HTML via Markdig
            Dim bodyHtml As String = ""
            If Not String.IsNullOrWhiteSpace(modelText) Then
                Try
                    Dim pipeline = New MarkdownPipelineBuilder().UseAdvancedExtensions().Build()
                    bodyHtml = Markdig.Markdown.ToHtml(Global.SharedLibrary.SharedLibrary.SharedMethods.NormalizeMarkdownForHtmlDisplay(modelText), pipeline)
                Catch
                    ' Fallback: plain-text HTML encoding
                    bodyHtml = "<p>" & System.Net.WebUtility.HtmlEncode(modelText).Replace(vbLf, "<br/>") & "</p>"
                End Try
            End If

            ' Build HTML page
            Dim html As New System.Text.StringBuilder()
            html.AppendLine("<!DOCTYPE html><html><head><meta charset=""utf-8"">")
            html.AppendLine("<style>")
            html.AppendLine("body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; font-size: 10pt; line-height: 1.5; padding: 20px; margin: 0; }")
            html.AppendLine("img { max-width: 100%; height: auto; border: 1px solid #ccc; margin-top: 10px; }")
            html.AppendLine("p { margin: 6px 0; }")
            html.AppendLine("ul, ol { margin-left: 20px; }")
            html.AppendLine("li { margin-bottom: 6px; }")
            html.AppendLine("h1, h2, h3 { color: #333; }")
            html.AppendLine("strong { color: #003366; }")
            html.AppendLine("code { background: #f6f8fa; padding: 2px 4px; border-radius: 3px; }")
            html.AppendLine("pre { background: #f6f8fa; padding: 10px; border-radius: 4px; overflow-x: auto; }")
            html.AppendLine(".path { font-size: 8pt; color: #666; word-break: break-all; margin-top: 8px; }")
            html.AppendLine("</style></head><body>")

            If bodyHtml.Length > 0 Then
                html.AppendLine(bodyHtml)
            End If

            If Not String.IsNullOrWhiteSpace(imagePath) AndAlso File.Exists(imagePath) Then
                ' Use file:// URI for local image display in WebBrowser control
                Dim fileUri As String = New Uri(imagePath).AbsoluteUri
                html.AppendLine("<img src=""" & fileUri & """ alt=""Generated Image"" />")
                html.AppendLine("<p class=""path"">Saved to: " & System.Net.WebUtility.HtmlEncode(imagePath) & "</p>")
            ElseIf Not String.IsNullOrWhiteSpace(imagePath) Then
                html.AppendLine("<p><em>Image file not found at: " & System.Net.WebUtility.HtmlEncode(imagePath) & "</em></p>")
            End If

            html.AppendLine("</body></html>")

            ' Build additional buttons — use fully qualified Action to avoid ambiguity
            ' with Microsoft.Office.Tools.Word.Action
            Dim capturedPath As String = imagePath
            Dim additionalButtons As New System.Collections.Generic.List(Of
                System.Tuple(Of String, System.Action, Boolean))()

            If Not String.IsNullOrWhiteSpace(capturedPath) AndAlso File.Exists(capturedPath) Then
                ' "Copy Path" button
                additionalButtons.Add(System.Tuple.Create(Of String, System.Action, Boolean)(
                    "Copy Path",
                    CType(Sub()
                              Try
                                  System.Windows.Forms.Clipboard.SetText(capturedPath)
                              Catch
                              End Try
                          End Sub, System.Action),
                    False))

                ' "Copy Image" button
                additionalButtons.Add(System.Tuple.Create(Of String, System.Action, Boolean)(
                    "Copy Image",
                    CType(Sub()
                              Try
                                  Using img = System.Drawing.Image.FromFile(capturedPath)
                                      System.Windows.Forms.Clipboard.SetImage(img)
                                  End Using
                              Catch ex As System.Exception
                                  ShowCustomMessageBox("Could not copy image to clipboard: " & ex.Message)
                              End Try
                          End Sub, System.Action),
                    False))

                ' "Insert into Document" button
                additionalButtons.Add(System.Tuple.Create(Of String, System.Action, Boolean)(
                    "Insert into Document",
                    CType(Sub()
                              Try
                                  InsertImageIntoActiveDocument(capturedPath)
                              Catch ex As System.Exception
                                  ShowCustomMessageBox("Error inserting image: " & ex.Message)
                              End Try
                          End Sub, System.Action),
                    True))
            End If

            ShowHTMLCustomMessageBox(
                html.ToString(),
                AN & " Image Generation",
                "",
                Nothing,
                Nothing,
                False,
                If(additionalButtons.Count > 0, additionalButtons.ToArray(), Nothing))

        Catch ex As System.Exception
            ' Ensure config is restored on any unexpected error
            If originalConfigLoaded Then
                Try : RestoreDefaults(_context, originalConfig) : Catch : End Try
                originalConfigLoaded = False
            End If
            ShowCustomMessageBox("Error in Image Generation: " & ex.Message)
        End Try
    End Sub


    ''' <summary>
    ''' Inserts an image file as an InlineShape at the current selection in the active Word document.
    ''' </summary>
    ''' <param name="imagePath">Full path to the image file to insert.</param>
    Private Sub InsertImageIntoActiveDocument(imagePath As String)
        If String.IsNullOrWhiteSpace(imagePath) OrElse Not File.Exists(imagePath) Then
            ShowCustomMessageBox("The image file could not be found.")
            Return
        End If

        Try
            Dim application As Word.Application = Globals.ThisAddIn.Application

            ' Ensure Word is activated and has focus before accessing the selection
            Try
                application.Activate()
                If application.Documents.Count > 0 Then
                    application.ActiveDocument.ActiveWindow.Activate()
                End If
            Catch
                ' Best-effort activation — continue even if it fails
            End Try

            Dim selection As Word.Selection = application.Selection
            Dim rng As Word.Range = selection.Range

            Dim targetRange As Object = rng
            rng.InlineShapes.AddPicture(
                FileName:=imagePath,
                LinkToFile:=False,
                SaveWithDocument:=True,
                Range:=targetRange)
        Catch ex As System.Exception
            ShowCustomMessageBox("Could not insert the image into the document: " & ex.Message)
        End Try
    End Sub

End Class