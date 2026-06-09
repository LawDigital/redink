' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.SpecialServices.vb
' Purpose: Executes user-selected special services by collecting parameters,
'          preparing prompts, invoking configured LLM endpoints, and displaying
'          the returned content inside Word.
'
' Architecture:
'  - Configuration Validation: Ensures INI data is loaded and the special service
'    path is available before proceeding.
'  - Parameter Parsing: Reads up to four INI-defined parameter strings, extracts
'    types/ranges/options, and builds SharedMethods.InputParameter definitions for
'    user input via ShowCustomVariableInputForm.
'  - Query Assistant (optional): Offers a helper prompt derived from the current
'    selection when SP_QueryPrompt is configured.
'  - Prompt Execution: Calls LLM with the collected prompt/selection data and
'    retains SP_MergePrompt state for downstream use.
'  - Result Presentation: Displays responses in a side pane or custom window,
'    supports markdown insertion into Word, and handles clipboard operations.
'  - Post-processing: Restores special text elements, updates fields, and reverts
'    configuration overrides.
'
' External Dependencies:
'  - SharedLibrary.SharedLibrary.SharedMethods for UI helpers, clipboard, markdown,
'    prompt interpolation, and LLM invocation.
'  - Microsoft.Office.Interop.Word for document and selection manipulation.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Data
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Drives the special service workflow by validating configuration, collecting optional parameters,
    ''' invoking the configured LLM, and presenting the generated content to the user.
    ''' </summary>
    Public Async Sub SpecialModel()
        Try
            If INILoadFail() Then Return
            Dim DoPane As Boolean = True
            Dim NoSelectedText As Boolean = False
            Dim AlternateText As String = ""
            Dim specialServicePath As String = ExpandEnvironmentVariables(INI_SpecialServicePath)

            If String.IsNullOrWhiteSpace(INI_SpecialServicePath) Then
                ShowCustomMessageBox("No special service path is configured.")
                Return
            End If

            If INILoadFail() Then Return
            Dim application As Word.Application = Globals.ThisAddIn.Application
            Dim selection As Word.Selection = application.Selection

            OptionChecked = False

            Debug.WriteLine($"SpecialModel invoked. Service path: '{specialServicePath}'")

            If Not ShowModelSelection(_context, specialServicePath, "Special Service", "Select the special service you want To query:", "Output in a pane (not directly in the document)", 2) Then
                originalConfigLoaded = False
                Return
            End If

            If selection.Type = Word.WdSelectionType.wdSelectionIP Then
                AlternateText = ShowCustomInputBox($"Provide the text for querying {INI_Model_2} (since you have not selected any text):", $"{AN} Special Service", False)
                NoSelectedText = True
                If AlternateText = "ESC" OrElse AlternateText.Trim() = "" Then
                    Return
                End If
            End If

            Dim iniValues() As String = {HideEscape(INI_Model_Parameter1), HideEscape(INI_Model_Parameter2), HideEscape(INI_Model_Parameter3), HideEscape(INI_Model_Parameter4)}
            Dim parameterDefs As New List(Of SharedLibrary.SharedLibrary.SharedMethods.InputParameter)()
            Dim typesList As New List(Of String)()
            Dim rangesList As New List(Of Tuple(Of Integer, Integer))()
            Dim optsDisplayList As New List(Of List(Of String))()
            Dim optsCodeList As New List(Of List(Of String))()

            For Each raw As String In iniValues
                If String.IsNullOrWhiteSpace(raw) Then Continue For

                Dim segments = raw.Split(";"c).Select(Function(s) s.Trim()).ToArray()
                If segments.Length = 0 Then Continue For

                Dim desc As String = segments(0)
                Dim t As String = If(segments.Length > 1 AndAlso Not String.IsNullOrEmpty(segments(1)), segments(1).ToLowerInvariant(), "string")
                Dim defaultStr As String = If(segments.Length > 2, segments(2), String.Empty)

                ' Range (numeric only) and options
                Dim rangeTuple As Tuple(Of Integer, Integer) = Nothing
                Dim optsRaw As List(Of String) = Nothing

                If (t = "integer" OrElse t = "long" OrElse t = "double") AndAlso segments.Length > 3 AndAlso System.Text.RegularExpressions.Regex.IsMatch(segments(3), "^\d+\s*-\s*\d+$") Then
                    Dim parts = segments(3).Split("-"c).Select(Function(s) s.Trim()).ToArray()
                    Dim minVal = Integer.Parse(parts(0), Globalization.CultureInfo.InvariantCulture)
                    Dim maxVal = Integer.Parse(parts(1), Globalization.CultureInfo.InvariantCulture)
                    rangeTuple = Tuple.Create(minVal, maxVal)

                    If segments.Length > 4 AndAlso Not String.IsNullOrWhiteSpace(segments(4)) Then
                        optsRaw = segments(4).Split(","c).Select(Function(o) o.Trim()).ToList()
                    End If
                End If

                If t = "string" AndAlso segments.Length > 3 AndAlso Not String.IsNullOrWhiteSpace(segments(3)) Then
                    optsRaw = segments(3).Split(","c).Select(Function(o) o.Trim()).ToList()
                End If

                ' Split options into display/code
                Dim displayList As List(Of String) = Nothing
                Dim codeList As List(Of String) = Nothing
                If optsRaw IsNot Nothing Then
                    displayList = New List(Of String)()
                    codeList = New List(Of String)()
                    For Each o As String In optsRaw
                        Dim lbl = o
                        Dim code = o
                        Dim idx1 = o.IndexOf("<"c)
                        Dim idx2 = o.IndexOf(">"c)
                        If idx1 >= 0 AndAlso idx2 > idx1 Then
                            lbl = o.Substring(0, idx1).Trim()
                            code = o.Substring(idx1 + 1, idx2 - idx1 - 1).Trim()
                        End If
                        displayList.Add(lbl)
                        codeList.Add(code)
                    Next
                End If

                ' Default for UI: prefer display label if code match exists
                Dim defaultDisplay As Object = defaultStr
                If codeList IsNot Nothing Then
                    Dim idxDef = codeList.IndexOf(defaultStr)
                    If idxDef >= 0 Then defaultDisplay = displayList(idxDef)
                End If

                ' Parse default by type
                Dim val As Object
                Select Case t
                    Case "boolean"
                        Dim b As Boolean
                        Boolean.TryParse(defaultStr, b)
                        val = b
                    Case "integer"
                        Dim i As Integer
                        Integer.TryParse(defaultStr, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, i)
                        val = i
                    Case "long"
                        Dim l As Long
                        Long.TryParse(defaultStr, Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, l)
                        val = l
                    Case "double"
                        Dim d As Double
                        ' Normalize comma to dot, then parse with invariant culture
                        Dim normalizedDefault As String = defaultStr.Replace(","c, "."c)
                        Double.TryParse(normalizedDefault, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, d)
                        val = d
                    Case Else
                        val = defaultDisplay
                End Select

                ' Build parameterDef
                If displayList IsNot Nothing Then
                    parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter(desc, val, displayList))
                Else
                    parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter(desc, val))
                End If

                ' Keep metadata aligned
                typesList.Add(t)
                rangesList.Add(rangeTuple)
                optsDisplayList.Add(displayList)
                optsCodeList.Add(codeList)
            Next

            Dim runQueryAssistant As Boolean = False

            If Not String.IsNullOrWhiteSpace(SP_QueryPrompt) Then
                parameterDefs.Add(New SharedLibrary.SharedLibrary.SharedMethods.InputParameter("Run query assistant", False))
                runQueryAssistant = True
            End If

            OtherPrompt = ""

            If parameterDefs.Count > 0 Then
                Dim parameters() As SharedLibrary.SharedLibrary.SharedMethods.InputParameter = parameterDefs.ToArray()
                If ShowCustomVariableInputForm("Please configure your parameters:", "Use '" & INI_Model_2 & "'", parameters) Then

                    ' Determine actual loop bounds robustly
                    Dim realParamCount As Integer = typesList.Count
                    Dim userParamCount As Integer = parameters.Length

                    ' Strip optional trailing boolean (query assistant) if present
                    If userParamCount > 0 AndAlso runQueryAssistant Then
                        Dim lastParam = parameters(userParamCount - 1)
                        If TypeOf lastParam.Value Is Boolean Then
                            runQueryAssistant = CType(lastParam.Value, Boolean)
                            userParamCount -= 1
                        Else
                            runQueryAssistant = False
                        End If
                    End If

                    Dim loopCount As Integer = System.Math.Min(userParamCount, realParamCount)

                    ' Read values with clamping/mapping
                    For i As Integer = 0 To loopCount - 1
                        Dim p As SharedLibrary.SharedLibrary.SharedMethods.InputParameter = parameters(i)
                        Dim rawValue As String = If(p.Value Is Nothing, String.Empty, p.Value.ToString()).Trim()
                        Dim t As String = typesList(i)
                        Dim range As Tuple(Of Integer, Integer) = rangesList(i)
                        Dim dispList = optsDisplayList(i)
                        Dim codeList = optsCodeList(i)

                        Dim paramValue As String

                        If TypeOf p.Value Is Boolean Then
                            paramValue = CType(p.Value, Boolean).ToString().ToLowerInvariant()
                        Else
                            ' Numeric clamping when applicable
                            If (t = "integer" OrElse t = "long" OrElse t = "double") AndAlso range IsNot Nothing Then
                                Dim num As Double
                                ' Normalize comma to dot, then parse with invariant culture
                                Dim normalizedValue As String = rawValue.Replace(","c, "."c)
                                If Double.TryParse(normalizedValue, Globalization.NumberStyles.Float, Globalization.CultureInfo.InvariantCulture, num) Then
                                    num = System.Math.Max(range.Item1, System.Math.Min(range.Item2, num))
                                    If t = "integer" OrElse t = "long" Then
                                        rawValue = CInt(System.Math.Round(num)).ToString(Globalization.CultureInfo.InvariantCulture)
                                    Else
                                        ' Always format with dot using invariant culture
                                        rawValue = num.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
                                    End If
                                End If
                            End If

                            ' Display → code mapping when list exists
                            If dispList IsNot Nothing Then
                                Dim idx As Integer = dispList.IndexOf(rawValue)
                                If idx >= 0 AndAlso idx < codeList.Count Then
                                    paramValue = UnHideEscape(codeList(idx))
                                Else
                                    paramValue = UnHideEscape(rawValue)
                                End If

                                Dim lower = paramValue.ToLowerInvariant()
                                If lower.StartsWith("(keine auswahl)") OrElse lower.StartsWith("(no selection)") OrElse paramValue.StartsWith("---") Then
                                    paramValue = ""
                                End If
                            Else
                                Dim rvLower As String = rawValue.ToLowerInvariant()
                                If rvLower.StartsWith("(keine auswahl)") OrElse rvLower.StartsWith("(no selection)") OrElse rawValue.StartsWith("---") Then
                                    rawValue = ""
                                End If
                                paramValue = UnHideEscape(rawValue)
                            End If
                        End If

                        ' Prompt passthrough
                        If p.Name IsNot Nothing AndAlso p.Name.ToLowerInvariant().Contains("prompt") Then
                            OtherPrompt = UnHideEscape(paramValue)
                        End If

                        ' Placeholder replacement
                        INI_Endpoint_2 = INI_Endpoint_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                        INI_APICall_2 = INI_APICall_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                        INI_APICall_Object_2 = INI_APICall_Object_2.Replace("{" & "parameter" & (i + 1) & "}", paramValue)
                    Next

                Else
                    Return
                End If
            End If

            If NoSelectedText Then
                SelectedText = AlternateText.Trim()
            Else
                SelectedText = selection.Text.Trim()
            End If

            If runQueryAssistant And Not NoSelectedText Then

                Dim querytext As String = Await LLM(SP_QueryPrompt, "<TEXTTOPROCESS>" & SelectedText & "</TEXTTOPROCESS>", "", "", 0, False)

                querytext = SLib.ShowCustomInputBox("This prompt has been generated based on your selection; modify it as you wish:", $"{AN} Query Assistant", False, querytext.Trim(), Context:=_context).Trim()
                If String.IsNullOrWhiteSpace(querytext) OrElse querytext.ToLower() = "esc" Then
                    Return
                End If
                SelectedText = querytext.Trim()

            End If

            ' ── Call LLM (or SSE transport for MCP servers) ──────────────
            Dim llmresult As String

            If INI_Endpoint_2.StartsWith(SharedMethods.MCP_SSE_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                Dim sseBase = INI_Endpoint_2.Substring(SharedMethods.MCP_SSE_PREFIX.Length)
                Dim resolvedHeaderB = INI_HeaderB_2.Replace("{apikey}", _context.DecodedAPI_2)

                Dim sseRequestBody As String = INI_APICall_2
                sseRequestBody = sseRequestBody.Replace("{promptuser}", SLib.CleanString(SelectedText))

                Dim splash As SharedLibrary.SharedLibrary.SplashScreenCountDown = Nothing

                Try
                    splash = New SharedLibrary.SharedLibrary.SplashScreenCountDown("Waiting for the AI to respond...", 0, 0, 0)
                    splash.Show()

                    llmresult = Await SharedMethods.ExecuteMCPSSEToolCall(
                        _context,
                        sseBase,
                        sseRequestBody,
                        INI_HeaderA_2,
                        resolvedHeaderB,
                        CInt(Math.Min(INI_Timeout_2, Integer.MaxValue)))
                Catch ex As Exception
                    ShowCustomMessageBox($"SSE tool call failed: {ex.Message}")
                    Return
                Finally
                    If splash IsNot Nothing Then
                        splash.Close()
                    End If
                End Try

                If Not String.IsNullOrWhiteSpace(llmresult) AndAlso
                   Not String.IsNullOrWhiteSpace(INI_Response_2) AndAlso
                   Not INI_Response_2.Equals("JSON", StringComparison.OrdinalIgnoreCase) Then
                    Try
                        Dim extracted = SharedLibrary.SharedLibrary.JsonTemplateFormatter.FormatJsonWithTemplate(llmresult, INI_Response_2)
                        If Not String.IsNullOrWhiteSpace(extracted) Then
                            llmresult = extracted
                        End If
                    Catch
                        ' If extraction fails, keep the raw JSON
                    End Try
                End If
            Else
                llmresult = Await LLM(OtherPrompt, SelectedText, "", "", 0, True)
            End If

            SP_MergePrompt_Cached = SP_MergePrompt

            If Not String.IsNullOrWhiteSpace(llmresult) Then
                If OptionChecked Then

                    Dim ClipPaneText1 As String = "Your service has provided the following result (you can edit it):"
                    Dim ClipText2 As String = "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting), or you can directly insert the original text in your document. If you select Cancel, nothing will be put into the clipboard."

                    If DoPane Then

                        If _uiContext IsNot Nothing Then  ' Make sure we run in the UI Thread
                            _uiContext.Post(Sub(s)

                                                ShowPaneAsync(
                                        ClipPaneText1,
                                        llmresult,
                                        "",
                                        AN,
                                        noRTF:=False,
                                        insertMarkdown:=True
                                        )
                                            End Sub, Nothing)
                        Else

                            ShowPaneAsync(ClipPaneText1, llmresult, "", AN, noRTF:=False, insertMarkdown:=True)

                        End If

                    Else

                        Dim dialogResult As String = ""

                        If _uiContext IsNot Nothing Then
                            Dim doneEvent As New ManualResetEventSlim(False)            ' Make sure we run in the UI Thread

                            _uiContext.Post(Sub(state)
                                                Try

                                                    Dim wordHwnd As IntPtr = GetWordMainWindowHandle()

                                                    dialogResult = ShowCustomWindow(ClipPaneText1,
                                                                            llmresult,
                                                                            ClipText2,
                                                                            AN,
                                                                            NoRTF:=False,
                                                                            Getfocus:=False,
                                                                            InsertMarkdown:=True,
                                                                            TransferToPane:=True,
                                                                            parentWindowHwnd:=wordHwnd)

                                                    If dialogResult <> "" And dialogResult <> "Pane" Then
                                                        If dialogResult = "Markdown" Then

                                                            Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                                            If NewDocChoice = 1 Then
                                                                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                                                Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                                                currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                InsertTextWithMarkdown(currentSelection, llmresult, True, True)
                                                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                                                If Regex.IsMatch(llmresult, pattern) Then
                                                                    Dim rng As Range = currentSelection.Range
                                                                    RestoreSpecialTextElements(rng)
                                                                    rng.Document.Fields.Update()
                                                                End If


                                                            ElseIf NewDocChoice = 2 Then
                                                                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                                                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                                                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & llmresult, False)
                                                                Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                                                If Regex.IsMatch(llmresult, pattern) Then
                                                                    Dim rng As Range = selection.Range
                                                                    RestoreSpecialTextElements(rng)
                                                                    rng.Document.Fields.Update()
                                                                End If

                                                            Else
                                                                ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                                                SLib.PutInClipboard(MarkdownToRtfConverter.Convert((llmresult)))
                                                            End If

                                                        Else
                                                            SLib.PutInClipboard(dialogResult)
                                                        End If
                                                    ElseIf dialogResult = "Pane" Then

                                                        ShowPaneAsync(
                                                                            ClipPaneText1,
                                                                            llmresult,
                                                                            "",
                                                                            AN,
                                                                            noRTF:=False,
                                                                            insertMarkdown:=True
                                                                            )
                                                    End If

                                                Finally
                                                    doneEvent.Set()
                                                End Try
                                            End Sub, Nothing)
                            ' doneEvent.Wait()

                        Else
                            dialogResult = ShowCustomWindow(
                                            ClipPaneText1,
                                            llmresult,
                                            ClipText2,
                                            AN,
                                            NoRTF:=False,
                                            Getfocus:=False,
                                            InsertMarkdown:=True,
                                            TransferToPane:=True)

                            If dialogResult <> "" And dialogResult <> "Pane" Then
                                If dialogResult = "Markdown" Then
                                    Dim NewDocChoice As Integer = ShowCustomYesNoBox("Do you want to insert the text into a new Word document (if you cancel, it will be in the clipboard with formatting)?", "Yes, new", "No, into my existing doc")

                                    If NewDocChoice = 1 Then
                                        Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                                        Dim currentSelection As Word.Selection = newDoc.Application.Selection
                                        currentSelection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        InsertTextWithMarkdown(currentSelection, llmresult, True, True)
                                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                        If Regex.IsMatch(llmresult, pattern) Then
                                            Dim rng As Range = currentSelection.Range
                                            RestoreSpecialTextElements(rng)
                                            rng.Document.Fields.Update()
                                        End If

                                    ElseIf NewDocChoice = 2 Then
                                        Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                        Globals.ThisAddIn.Application.Selection.TypeParagraph()
                                        InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & llmresult, False)
                                        Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                                        If Regex.IsMatch(llmresult, pattern) Then
                                            Dim rng As Range = wordApp.Selection.Range
                                            RestoreSpecialTextElements(rng)
                                            rng.Document.Fields.Update()
                                        End If

                                    Else
                                        ShowCustomMessageBox("No text was inserted (but included in the clipboard as RTF).")
                                        SLib.PutInClipboard(MarkdownToRtfConverter.Convert((llmresult)))
                                    End If
                                Else
                                    SLib.PutInClipboard(dialogResult)
                                End If
                            ElseIf dialogResult = "Pane" Then

                                ShowPaneAsync(
                                                    ClipPaneText1,
                                                    llmresult,
                                                    "",
                                                    AN,
                                                    noRTF:=False,
                                                    insertMarkdown:=True
                                                    )
                            End If

                        End If

                    End If
                Else
                    Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    Globals.ThisAddIn.Application.Selection.TypeParagraph()
                    InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, llmresult, False)
                    Dim pattern As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                    If Regex.IsMatch(llmresult, pattern) Then
                        Dim rng As Range = Globals.ThisAddIn.Application.Selection.Range
                        RestoreSpecialTextElements(rng)
                        rng.Document.Fields.Update()
                    End If
                End If
            End If

        Catch ex As System.Exception
            MessageBox.Show("Error in SpecialModel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            If originalConfig IsNot Nothing Then
                RestoreDefaults(_context, originalConfig)
            End If
            originalConfigLoaded = False
        End Try
    End Sub

End Class
