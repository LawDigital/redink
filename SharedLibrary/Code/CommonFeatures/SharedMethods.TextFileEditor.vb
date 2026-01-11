' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: SharedMethods.TextFileEditor.vb
' Purpose: Provides a modal WinForms-based text editor for an arbitrary file path,
'          including optional JSON formatting support and (when a context is provided)
'          an AI-powered JSON check that shows the result as rendered HTML.
'
' Behavior:
'  - Loads existing file contents (UTF-8 preferred; falls back to default encoding).
'  - Saves content as UTF-8 with BOM and creates a ".bak" backup when the target file exists.
'  - For JSON files (or when ForceJson=True), supports pretty/minified display toggling.
'    * Full-document mode: parses the entire document as JSON and toggles formatting.
'    * Embedded-segment mode (ForceJson=True only): finds JSON object/array segments in mixed text,
'      parses them, and replaces them with formatted JSON.
'  - Keyboard shortcuts:
'    * Ctrl+S: Save
'    * Ctrl+Shift+F: Toggle JSON Pretty/Minify (when available)
'
' Key internal helpers:
'  - ExtractEmbeddedJsonSegments: finds parseable JSON segments (object/array) within text.
'  - ReplaceSegmentsWithFormatting: replaces segments with a chosen JSON formatting style.
'  - TryFindBalancedJson: scans for balanced JSON object/array boundaries while honoring strings/escapes.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Windows.Forms
Imports Markdig
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        ''' <summary>
        ''' Shows a modal text editor for the specified file path, with optional JSON formatting support.
        ''' </summary>
        ''' <param name="filePath">Target file path to load and save.</param>
        ''' <param name="headerText">Text displayed above the editor.</param>
        ''' <param name="ForceJson">
        ''' If <c>True</c>, enables JSON formatting support even if the file extension is not <c>.json</c>.
        ''' In ForceJson mode, embedded JSON segments may be pretty-printed/minified within mixed text.
        ''' </param>
        ''' <param name="_context">
        ''' Optional shared context. If provided and JSON mode is active, an additional button "Check JSON with AI"
        ''' is shown to run an LLM check and display the result.
        ''' </param>
        Public Shared Sub ShowTextFileEditor(ByVal filePath As System.String, ByVal headerText As System.String, Optional ForceJson As Boolean = False, Optional _context As ISharedContext = Nothing, Optional ByRef wasSaved As System.Boolean? = Nothing)
            ' --- Guard & Input Validation ---
            Try
                If filePath Is Nothing OrElse filePath.Trim().Length = 0 Then
                    ShowCustomMessageBox("No file path was provided.")
                    Return
                End If
            Catch ex As System.Exception
                ShowCustomMessageBox("Unexpected error while validating input: " & ex.Message)
                Return
            End Try

            Dim localWasSaved As Boolean? = Nothing

            ' --- Create Form & Controls ---
            Dim editorForm As New System.Windows.Forms.Form()
            editorForm.Text = "Text File Editor"
            editorForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            editorForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Sizable
            editorForm.MinimizeBox = True
            editorForm.MaximizeBox = True
            editorForm.ShowInTaskbar = True
            editorForm.KeyPreview = True
            editorForm.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi

            ' Initial size based on screen (height = 60% of working area; width keeps 9:6 ratio)
            Try
                Dim scr As System.Windows.Forms.Screen = System.Windows.Forms.Screen.FromPoint(System.Windows.Forms.Cursor.Position)
                Dim wa As System.Drawing.Rectangle = scr.WorkingArea

                Dim targetHeight As System.Int32 = System.Convert.ToInt32(System.Math.Floor(wa.Height * 0.6R))
                If targetHeight < 540 Then targetHeight = 540

                Dim targetWidth As System.Int32 = System.Convert.ToInt32(System.Math.Floor(targetHeight * 9.0R / 6.0R))
                If targetWidth > wa.Width Then
                    targetWidth = wa.Width
                    targetHeight = System.Convert.ToInt32(System.Math.Floor(targetWidth * 6.0R / 9.0R))
                End If

                editorForm.ClientSize = New System.Drawing.Size(targetWidth, targetHeight)
                Dim minW As System.Int32 = System.Math.Max(780, System.Convert.ToInt32(System.Math.Floor(targetWidth / 2.0R)))
                Dim minH As System.Int32 = System.Math.Max(540, System.Convert.ToInt32(System.Math.Floor(targetHeight / 2.0R)))
                editorForm.MinimumSize = New System.Drawing.Size(minW, minH)
            Catch ex As System.Exception
                editorForm.ClientSize = New System.Drawing.Size(1560, 1080)
                editorForm.MinimumSize = New System.Drawing.Size(780, 540)
            End Try

            ' Set icon
            Try
                Dim bmp As New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
                editorForm.Icon = System.Drawing.Icon.FromHandle(bmp.GetHicon())
            Catch ex As System.Exception
                ' Non-fatal
            End Try

            ' Set predefined font
            Try
                Dim standardFont As New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
                editorForm.Font = standardFont
            Catch ex As System.Exception
                ' Non-fatal
            End Try

            ' Root container
            Dim rootPanel As New System.Windows.Forms.TableLayoutPanel()
            rootPanel.Dock = System.Windows.Forms.DockStyle.Fill
            rootPanel.BackColor = System.Drawing.Color.Transparent
            rootPanel.ColumnCount = 1
            rootPanel.RowCount = 3
            rootPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)) ' Label
            rootPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F)) ' Editor
            rootPanel.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize)) ' Buttons
            rootPanel.Padding = New System.Windows.Forms.Padding(15, 12, 15, 10)
            editorForm.Controls.Add(rootPanel)

            ' Header label
            Dim headerLabel As New System.Windows.Forms.Label()
            headerLabel.AutoSize = True
            headerLabel.Text = If(headerText, System.String.Empty)
            headerLabel.UseCompatibleTextRendering = True
            headerLabel.Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            headerLabel.MaximumSize = New System.Drawing.Size(editorForm.ClientSize.Width - (rootPanel.Padding.Left + rootPanel.Padding.Right), 0)
            headerLabel.Anchor = System.Windows.Forms.AnchorStyles.Left Or System.Windows.Forms.AnchorStyles.Right Or System.Windows.Forms.AnchorStyles.Top
            rootPanel.Controls.Add(headerLabel, 0, 0)

            ' Text editor
            Dim textEditor As New System.Windows.Forms.TextBox()
            textEditor.Multiline = True
            textEditor.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            textEditor.WordWrap = True
            textEditor.AcceptsReturn = True
            textEditor.AcceptsTab = True
            textEditor.Dock = System.Windows.Forms.DockStyle.Fill
            textEditor.Margin = New System.Windows.Forms.Padding(0, 0, 0, 8)
            textEditor.Font = New System.Drawing.Font("Consolas", 10.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            rootPanel.Controls.Add(textEditor, 0, 1)

            ' Bottom buttons
            Dim flowButtons As New System.Windows.Forms.FlowLayoutPanel()
            flowButtons.AutoSize = True
            flowButtons.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            flowButtons.FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight
            flowButtons.WrapContents = False
            flowButtons.Dock = System.Windows.Forms.DockStyle.Left
            flowButtons.Margin = New System.Windows.Forms.Padding(15, 15, 0, 15)
            flowButtons.Padding = New System.Windows.Forms.Padding(0)
            rootPanel.Controls.Add(flowButtons, 0, 2)

            ' Save button
            Dim btnSave As New System.Windows.Forms.Button()
            btnSave.Text = "&Save"
            btnSave.AutoSize = True
            btnSave.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            btnSave.Margin = New System.Windows.Forms.Padding(0, 0, 12, 0)
            btnSave.Padding = New System.Windows.Forms.Padding(5)

            ' Cancel button
            Dim btnCancel As New System.Windows.Forms.Button()
            btnCancel.Text = "Cancel"
            btnCancel.AutoSize = True
            btnCancel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
            btnCancel.Margin = New System.Windows.Forms.Padding(0)
            btnCancel.Padding = New System.Windows.Forms.Padding(5)

            ' JSON toggle button (added only when applicable later)
            Dim btnToggleJson As System.Windows.Forms.Button = Nothing

            flowButtons.Controls.Add(btnSave)
            flowButtons.Controls.Add(btnCancel)

            ' Enter = Save, Esc = Cancel
            editorForm.AcceptButton = btnSave
            editorForm.CancelButton = btnCancel

            ' Adjust label wrapping on resize
            AddHandler editorForm.Resize, Sub(sender As System.Object, e As System.EventArgs)
                                              Try
                                                  headerLabel.MaximumSize = New System.Drawing.Size(editorForm.ClientSize.Width - (rootPanel.Padding.Left + rootPanel.Padding.Right), 0)
                                              Catch ex As System.Exception
                                              End Try
                                          End Sub

            ' --- Load file content ---
            Dim originalLoadedText As String = System.String.Empty
            Try
                If System.IO.File.Exists(filePath) Then
                    Try
                        originalLoadedText = System.IO.File.ReadAllText(filePath, System.Text.Encoding.UTF8)
                    Catch exUtf8 As System.Exception
                        Try
                            originalLoadedText = System.IO.File.ReadAllText(filePath)
                        Catch exDefault As System.Exception
                            ShowCustomMessageBox("Failed to read file:" & System.Environment.NewLine & exDefault.Message)
                        End Try
                    End Try
                End If
            Catch ex As System.Exception
                ShowCustomMessageBox("Unexpected error while loading the file:" & System.Environment.NewLine & ex.Message)
            End Try
            textEditor.Text = originalLoadedText

            ' --- Minimal invasive JSON pretty-print support ---
            Dim isJsonFile As Boolean = False
            Dim originalJsonRaw As String = Nothing
            Dim formattedJson As String = Nothing
            Dim jsonCurrentlyFormatted As Boolean = False

            Try
                If filePath.EndsWith(".json", System.StringComparison.OrdinalIgnoreCase) Or ForceJson Then
                    isJsonFile = True

                    Dim trimmed = originalLoadedText.TrimStart()
                    Dim tok As Newtonsoft.Json.Linq.JToken = Nothing

                    ' Try full-document parse first (original behavior)
                    If trimmed.StartsWith("{"c) OrElse trimmed.StartsWith("["c) Then
                        Try
                            tok = Newtonsoft.Json.Linq.JToken.Parse(originalLoadedText)
                        Catch
                            tok = Nothing ' fall through to embedded-segment handling
                        End Try
                    End If

                    If tok IsNot Nothing Then
                        ' Full JSON document – keep existing pretty/minify toggle behavior
                        originalJsonRaw = originalLoadedText
                        formattedJson = tok.ToString(Newtonsoft.Json.Formatting.Indented)
                        textEditor.Text = formattedJson
                        jsonCurrentlyFormatted = True

                        btnToggleJson = New System.Windows.Forms.Button()
                        btnToggleJson.Text = "Minify JSON"
                        btnToggleJson.AutoSize = True
                        btnToggleJson.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                        btnToggleJson.Margin = New System.Windows.Forms.Padding(0, 0, 12, 0)
                        btnToggleJson.Padding = New System.Windows.Forms.Padding(5)

                        flowButtons.Controls.Add(btnToggleJson)
                        flowButtons.Controls.SetChildIndex(btnToggleJson, 1)

                        ' Toggles the editor text between indented and minified JSON.
                        AddHandler btnToggleJson.Click,
                            Sub()
                                Try
                                    If Not isJsonFile OrElse originalJsonRaw Is Nothing Then Return
                                    Dim tok2 = Newtonsoft.Json.Linq.JToken.Parse(originalJsonRaw)
                                    If jsonCurrentlyFormatted Then
                                        textEditor.Text = tok2.ToString(Newtonsoft.Json.Formatting.None)
                                        btnToggleJson.Text = "Pretty JSON"
                                    Else
                                        textEditor.Text = tok2.ToString(Newtonsoft.Json.Formatting.Indented)
                                        btnToggleJson.Text = "Minify JSON"
                                    End If
                                    jsonCurrentlyFormatted = Not jsonCurrentlyFormatted
                                Catch toggleEx As System.Exception
                                    ShowCustomMessageBox("JSON toggle failed: " & toggleEx.Message)
                                End Try
                            End Sub

                    ElseIf ForceJson Then
                        ' Mixed text: find and pretty-print only embedded JSON segments
                        Dim segments = SharedMethods.ExtractEmbeddedJsonSegments(originalLoadedText, requireObjectInArray:=True)

                        If segments IsNot Nothing AndAlso segments.Count > 0 Then
                            originalJsonRaw = originalLoadedText
                            Dim prettyText = SharedMethods.ReplaceSegmentsWithFormatting(originalJsonRaw, segments, Newtonsoft.Json.Formatting.Indented)
                            Dim minText = SharedMethods.ReplaceSegmentsWithFormatting(originalJsonRaw, segments, Newtonsoft.Json.Formatting.None)

                            textEditor.Text = prettyText
                            jsonCurrentlyFormatted = True

                            btnToggleJson = New System.Windows.Forms.Button()
                            btnToggleJson.Text = "Minify JSON"
                            btnToggleJson.AutoSize = True
                            btnToggleJson.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                            btnToggleJson.Margin = New System.Windows.Forms.Padding(0, 0, 12, 0)
                            btnToggleJson.Padding = New System.Windows.Forms.Padding(5)

                            flowButtons.Controls.Add(btnToggleJson)
                            flowButtons.Controls.SetChildIndex(btnToggleJson, 1)

                            ' Toggles the editor text between pretty-printed and minified embedded JSON segments.
                            AddHandler btnToggleJson.Click,
                                Sub()
                                    Try
                                        If Not isJsonFile OrElse originalJsonRaw Is Nothing Then Return
                                        If jsonCurrentlyFormatted Then
                                            textEditor.Text = minText
                                            btnToggleJson.Text = "Pretty JSON"
                                        Else
                                            textEditor.Text = prettyText
                                            btnToggleJson.Text = "Minify JSON"
                                        End If
                                        jsonCurrentlyFormatted = Not jsonCurrentlyFormatted
                                    Catch toggleEx As System.Exception
                                        ShowCustomMessageBox("JSON toggle failed: " & toggleEx.Message)
                                    End Try
                                End Sub
                        End If
                    End If

                    If _context IsNot Nothing Then
                        ' LLM Process button
                        Dim btnProcessLLM As New System.Windows.Forms.Button()
                        btnProcessLLM.Text = "Check JSON with AI"
                        btnProcessLLM.AutoSize = True
                        btnProcessLLM.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                        btnProcessLLM.Margin = New System.Windows.Forms.Padding(0, 0, 12, 0)
                        btnProcessLLM.Padding = New System.Windows.Forms.Padding(5)

                        flowButtons.Controls.Add(btnProcessLLM)
                        ' Insert after toggle button if it exists, otherwise insert at position 1
                        If btnToggleJson IsNot Nothing Then
                            flowButtons.Controls.SetChildIndex(btnProcessLLM, 2)
                        Else
                            flowButtons.Controls.SetChildIndex(btnProcessLLM, 1)
                        End If

                        ' Checks the current selection (or full editor content) using the LLM and shows the result as HTML.
                        AddHandler btnProcessLLM.Click,
                        Async Sub(sender As Object, e As EventArgs)
                            Try
                                btnProcessLLM.Enabled = False
                                btnProcessLLM.Text = "Processing..."

                                ' Get text to process - selected text or entire content
                                Dim textToProcess As String
                                If textEditor.SelectionLength > 0 Then
                                    textToProcess = textEditor.SelectedText
                                Else
                                    textToProcess = textEditor.Text
                                End If

                                If String.IsNullOrWhiteSpace(textToProcess) Then
                                    ShowCustomMessageBox("No text to process.", AN)
                                    Return
                                End If

                                ' Define system prompt
                                Dim systemPrompt As String = JSONCheckPrompt

                                ' Call LLM function
                                Dim result As String = Await LLM(_context, systemPrompt, textToProcess)

                                If String.IsNullOrWhiteSpace(result) Then
                                    ShowCustomMessageBox("The AI did not provide any response (which is an error).")
                                Else
                                    ' Convert Markdown to HTML using Markdig
                                    Dim pipeline = New MarkdownPipelineBuilder() _
                                    .UseAdvancedExtensions() _
                                    .UseEmojiAndSmiley() _
                                    .UseEmphasisExtras() _
                                    .UseFootnotes() _
                                    .UsePipeTables() _
                                    .UseTaskLists() _
                                    .UseAutoLinks() _
                                    .Build()

                                    Dim htmlResult As String = Markdown.ToHtml(result, pipeline)

                                    ' Display result
                                    ShowHTMLCustomMessageBox(htmlResult, "AI Analysis Result")
                                End If

                            Catch ex As Exception
                                ShowCustomMessageBox("Error processing with AI: " & ex.Message, AN)
                            Finally
                                btnProcessLLM.Enabled = True
                                btnProcessLLM.Text = "Check JSON with AI"
                            End Try
                        End Sub
                    End If

                End If
            Catch ex As System.Exception
                ' Non-fatal JSON formatting failure
            End Try

            ' Save logic
            ''' <summary>
            ''' Saves the current editor content to disk and creates a ".bak" backup when the target file exists.
            ''' </summary>
            Dim doSave As System.Action =
        Sub()
            Try
                Dim dir As System.String = System.IO.Path.GetDirectoryName(filePath)
                If dir Is Nothing OrElse dir.Trim().Length = 0 Then
                    ShowCustomMessageBox("Invalid file path or directory.")
                    Return
                End If
                If Not System.IO.Directory.Exists(dir) Then
                    ShowCustomMessageBox("Directory does not exist: " & dir)
                    Return
                End If

                Dim bakPath As System.String = filePath & ".bak"

                If System.IO.File.Exists(filePath) Then
                    Try
                        System.IO.File.Copy(filePath, bakPath, True)
                    Catch exCopy As System.Exception
                        ShowCustomMessageBox("Failed to create backup file:" & System.Environment.NewLine & exCopy.Message)
                        Return
                    End Try
                End If

                Try
                    Dim enc As System.Text.Encoding = New System.Text.UTF8Encoding(True)
                    System.IO.File.WriteAllText(filePath, textEditor.Text, enc)
                Catch exWrite As System.Exception
                    ShowCustomMessageBox("Failed to save file:" & System.Environment.NewLine & exWrite.Message)
                    Return
                End Try

                localWasSaved = True
                editorForm.DialogResult = System.Windows.Forms.DialogResult.OK
                editorForm.Close()

            Catch ex As System.Exception
                ShowCustomMessageBox("Unexpected error while saving:" & System.Environment.NewLine & ex.Message)
            End Try
        End Sub

            ' Event bindings
            AddHandler btnSave.Click, Sub(sender As System.Object, e As System.EventArgs)
                                          doSave()
                                      End Sub

            AddHandler btnCancel.Click, Sub(sender As System.Object, e As System.EventArgs)
                                            localWasSaved = False
                                            editorForm.DialogResult = System.Windows.Forms.DialogResult.Cancel
                                            editorForm.Close()
                                        End Sub

            AddHandler editorForm.FormClosing,
                                        Sub(sender As Object, e As System.Windows.Forms.FormClosingEventArgs)
                                            If Not localWasSaved.HasValue Then
                                                localWasSaved = False
                                            End If
                                        End Sub

            ' Keyboard shortcuts: Ctrl+S (save), Ctrl+Shift+F (toggle JSON formatting)
            AddHandler editorForm.KeyDown,
        Sub(sender As System.Object, e As System.Windows.Forms.KeyEventArgs)
            Try
                If e.Control AndAlso e.KeyCode = System.Windows.Forms.Keys.S Then
                    e.SuppressKeyPress = True
                    doSave()
                ElseIf e.Control AndAlso e.Shift AndAlso e.KeyCode = System.Windows.Forms.Keys.F Then
                    If btnToggleJson IsNot Nothing Then
                        e.SuppressKeyPress = True
                        btnToggleJson.PerformClick()
                    End If
                End If
            Catch ex As System.Exception
            End Try
        End Sub

            ' Sets the initial cursor/selection when the form is shown.
            AddHandler editorForm.Shown,
                Sub(sender As System.Object, e As System.EventArgs)
                    Try
                        textEditor.SelectionStart = 0
                        textEditor.SelectionLength = 0
                    Catch ex As System.Exception
                    End Try
                End Sub

            ' Show modal window
            Try
                Dim active As System.Windows.Forms.IWin32Window = System.Windows.Forms.Form.ActiveForm
                If active IsNot Nothing Then
                    editorForm.ShowDialog(active)
                Else
                    editorForm.ShowDialog()
                End If
            Catch ex As System.Exception
                Try
                    editorForm.Show()
                Catch exShow As System.Exception
                    ShowCustomMessageBox("Failed to display editor window:" & System.Environment.NewLine & exShow.Message)
                End Try
            End Try

            wasSaved = localWasSaved

        End Sub

        ' Keep the segment type non-public
        ''' <summary>
        ''' Represents a JSON segment found within a larger text, including its position and parsed token.
        ''' </summary>
        Private Structure JsonSegment
            ''' <summary>
            ''' Start offset of the segment in the original text.
            ''' </summary>
            Public Start As Integer

            ''' <summary>
            ''' Length of the segment in characters.
            ''' </summary>
            Public Length As Integer

            ''' <summary>
            ''' Parsed JSON token for this segment.
            ''' </summary>
            Public Token As Newtonsoft.Json.Linq.JToken
        End Structure

        ' Internal-only: do not expose private type publicly
        ''' <summary>
        ''' Scans a text for JSON object/array segments and returns those that can be parsed as JSON.
        ''' </summary>
        ''' <param name="text">Source text to scan.</param>
        ''' <param name="requireObjectInArray">
        ''' If <c>True</c>, skips array segments that do not contain a '{' character (used to ignore bracketed headings like "[Something]").
        ''' </param>
        ''' <returns>List of parsed JSON segments (may be empty).</returns>
        Private Shared Function ExtractEmbeddedJsonSegments(text As String,
                                                   Optional requireObjectInArray As Boolean = True) _
                                                   As List(Of JsonSegment)
            Dim result As New List(Of JsonSegment)()
            If String.IsNullOrEmpty(text) Then Return result

            Dim i As Integer = 0
            While i < text.Length
                Dim ch = text(i)
                If ch <> "{"c AndAlso ch <> "["c Then
                    i += 1
                    Continue While
                End If

                Dim endIdx As Integer
                If TryFindBalancedJson(text, i, endIdx) Then
                    Dim segment = text.Substring(i, endIdx - i + 1)

                    ' Skip bracketed headings like [Something] (no JSON objects inside)
                    If ch = "["c AndAlso requireObjectInArray AndAlso segment.IndexOf("{"c) = -1 Then
                        i += 1
                        Continue While
                    End If

                    Try
                        Dim tok = Newtonsoft.Json.Linq.JToken.Parse(segment)
                        result.Add(New JsonSegment With {.Start = i, .Length = endIdx - i + 1, .Token = tok})
                        i = endIdx + 1
                        Continue While
                    Catch
                        ' Not valid JSON – advance
                    End Try
                End If

                i += 1
            End While

            Return result
        End Function

        ' Internal-only: do not expose private type publicly
        ''' <summary>
        ''' Replaces the specified JSON segments in the input text with the JSON token string using the specified formatting.
        ''' </summary>
        ''' <param name="text">Original text.</param>
        ''' <param name="segments">JSON segments to replace.</param>
        ''' <param name="fmt">JSON formatting mode used to render the tokens.</param>
        ''' <returns>Text with segments replaced, or the original text when no segments are provided.</returns>
        Private Shared Function ReplaceSegmentsWithFormatting(text As String,
                                                     segments As List(Of JsonSegment),
                                                     fmt As Newtonsoft.Json.Formatting) As String
            If segments Is Nothing OrElse segments.Count = 0 Then Return text

            segments = segments.OrderBy(Function(s) s.Start).ToList()

            Dim sb As New System.Text.StringBuilder(text.Length + 1024)
            Dim last As Integer = 0
            For Each seg In segments
                If seg.Start > last Then
                    sb.Append(text, last, seg.Start - last)
                End If
                sb.Append(seg.Token.ToString(fmt))
                last = seg.Start + seg.Length
            Next
            If last < text.Length Then sb.Append(text, last, text.Length - last)

            Return sb.ToString()
        End Function

        ' Balanced scanner that honors strings and escapes to find end index of a JSON object/array
        ''' <summary>
        ''' Attempts to find the end index of a JSON object/array starting at <paramref name="startIndex"/>,
        ''' balancing curly/square brackets while honoring JSON string escaping rules.
        ''' </summary>
        ''' <param name="s">Source text.</param>
        ''' <param name="startIndex">Index of the initial '{' or '[' character.</param>
        ''' <param name="endIndex">On success, receives the index of the matching closing brace/bracket.</param>
        ''' <returns><c>True</c> if a balanced end was found; otherwise <c>False</c>.</returns>
        Private Shared Function TryFindBalancedJson(s As String, startIndex As Integer, ByRef endIndex As Integer) As Boolean
            If startIndex < 0 OrElse startIndex >= s.Length Then Return False

            Dim startCh = s(startIndex)
            If startCh <> "{"c AndAlso startCh <> "["c Then Return False

            Dim inString As Boolean = False
            Dim escapeNext As Boolean = False
            Dim curly As Integer = If(startCh = "{"c, 1, 0)
            Dim square As Integer = If(startCh = "["c, 1, 0)

            For i = startIndex + 1 To s.Length - 1
                Dim ch = s(i)

                If inString Then
                    If escapeNext Then
                        escapeNext = False
                    ElseIf ch = "\"c Then
                        escapeNext = True
                    ElseIf ch = """"c Then
                        inString = False
                    End If
                    Continue For
                End If

                If ch = """"c Then
                    inString = True
                ElseIf ch = "{"c Then
                    curly += 1
                ElseIf ch = "}"c Then
                    curly -= 1
                    If curly < 0 Then Return False
                ElseIf ch = "["c Then
                    square += 1
                ElseIf ch = "]"c Then
                    square -= 1
                    If square < 0 Then Return False
                End If

                If curly = 0 AndAlso square = 0 Then
                    endIndex = i
                    Return True
                End If
            Next

            Return False
        End Function
    End Class
End Namespace