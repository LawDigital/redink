' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.DocCheck.vb
' Purpose: Coordinates the DocCheck workflow for Microsoft Word, including text
'          acquisition, rule-set management, LLM invocations, and presentation
'          of findings as Word comments or panes.
'
' Architecture:
'  - Configuration: Expands and validates DocCheck paths, manages optional
'    secondary model loading, and honors INI-driven defaults.
'  - Input Acquisition: Captures text from the selection, entire document, or
'    external files and gathers optional supporting documents.
'  - Rule Set Management: Enumerates *.txt rule files, parses metadata
'    (prompts, notices, Markdown overrides), and builds DocCheckRuleSet objects.
'  - Execution Flow: Collects runtime parameters, runs isolated or multi-clause
'    analyses, and streams record prompts to the LLM with cancellation support.
'  - Output Handling: Applies comment bubbles, summaries, notices, and UI panes
'    while guarding Word selections and respecting Markdown settings.
'  - Utilities: Provides JSON extraction, display-name uniqueness, and selection
'    redirection helpers to keep the process stable for reviewers.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

''' <summary>
''' Provides DocCheck orchestration for the Word add-in, covering selection capture, rule-set handling, and LLM integration.
''' </summary>
Partial Public Class ThisAddIn

    ''' <summary>
    ''' Asynchronously runs the DocCheck workflow from text gathering through rule-set execution and output presentation.
    ''' </summary>
    Public Async Sub RunDocCheck()

        If INILoadFail() OrElse Not IsDocumentEditable() Then Return

        ' ── Mode selector: Standard DocCheck or Markup Review ──
        Dim modeAnswer As System.Int32 = ShowCustomYesNoBox(
            "Which type of document check would you like to run?" & vbCrLf & vbCrLf &
            "• Check Requirements — Analyzes the document (or a selection) against a " &
            "rule set of compliance criteria (from your DocCheck script files). " &
            "Each rule is checked independently and findings are reported as " &
            "Word comments or a summary report." & vbCrLf & vbCrLf &
            "• Markup Review — Compares tracked changes in the document against " &
            "acceptability constraints defined in a separate playbook (.docx with " &
            "comments). Evaluates whether each revision is within acceptable bounds, " &
            "suggests compromise redrafts where needed, and optionally checks " &
            "whether other clauses undermine the constraints.",
            "Check Requirements",
            "Markup Review",
            AN & " Document Check")

        If modeAnswer = 0 Then
            Return
        ElseIf modeAnswer = 2 Then
            Await RunMarkupReview()
            Return
        End If
        ' modeAnswer = 1 → proceed with standard DocCheck below

        ' Expand and normalize the configured paths
        Dim DocCheckPath As System.String = ExpandEnvironmentVariables(INI_DocCheckPath)
        If Not System.String.IsNullOrEmpty(DocCheckPath) AndAlso Not DocCheckPath.EndsWith("\", System.StringComparison.Ordinal) Then
            DocCheckPath &= "\"
        End If

        Dim DocCheckPathLocal As System.String = ExpandEnvironmentVariables(INI_DocCheckPathLocal)
        If Not System.String.IsNullOrEmpty(DocCheckPathLocal) AndAlso Not DocCheckPathLocal.EndsWith("\", System.StringComparison.Ordinal) Then
            DocCheckPathLocal &= "\"
        End If

        Dim do2ndModel As System.Boolean = False

        Try
            ' 0) Validate paths (at least one must be defined and exist)
            Dim hasGlobal As System.Boolean = (DocCheckPath IsNot Nothing AndAlso DocCheckPath.Trim().Length > 0 AndAlso System.IO.Directory.Exists(DocCheckPath) = True)
            Dim hasLocal As System.Boolean = (DocCheckPathLocal IsNot Nothing AndAlso DocCheckPathLocal.Trim().Length > 0 AndAlso System.IO.Directory.Exists(DocCheckPathLocal) = True)

            If hasGlobal = False AndAlso hasLocal = False Then
                ShowCustomMessageBox("No DocCheck paths are configured or accessible. Please configure at least one path (using 'DocCheckPath' or 'DocCheckPathLocal').")
                Return
            End If

            ' 1) Acquire text (selection → otherwise whole document), but KEEP a Selection (not just a Range)
            Dim app As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            If app Is Nothing Then
                ShowCustomMessageBox("Word application was not found.")
                Return
            End If
            If app.Documents Is Nothing OrElse app.Documents.Count = 0 Then
                ShowCustomMessageBox("No open document.")
                Return
            End If

            Dim doc As Microsoft.Office.Interop.Word.Document = app.ActiveDocument
            If doc Is Nothing Then
                ShowCustomMessageBox("Active document was not found.")
                Return
            End If

            Dim currentSelection As Microsoft.Office.Interop.Word.Selection = app.Selection
            Dim FromFile As String = ""

            If currentSelection.Type = WdSelectionType.wdSelectionIP Then

                Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you instead want to analyze text from a document file or Powerpoint presentation?", "Yes", "No, proceed with this text", AN & " Document Check", Nothing, "", "Edit Script Files",
                                                                                                                       extraButtonAction:=Sub()
                                                                                                                                              Try
                                                                                                                                                  ' Collect DocCheck script files from both paths
                                                                                                                                                  Dim displayToPath As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                                                                                                                                  Dim options As New List(Of String)

                                                                                                                                                  Dim dcPaths As New List(Of (p As String, isLocal As Boolean))
                                                                                                                                                  If Not String.IsNullOrWhiteSpace(DocCheckPath) Then
                                                                                                                                                      dcPaths.Add((DocCheckPath, False))
                                                                                                                                                  End If
                                                                                                                                                  If Not String.IsNullOrWhiteSpace(DocCheckPathLocal) Then
                                                                                                                                                      dcPaths.Add((DocCheckPathLocal, True))
                                                                                                                                                  End If

                                                                                                                                                  For Each tuple In dcPaths
                                                                                                                                                      Dim basePath = tuple.p
                                                                                                                                                      Dim isLocal = tuple.isLocal
                                                                                                                                                      If IO.Directory.Exists(basePath) Then
                                                                                                                                                          Dim files = IO.Directory.GetFiles(basePath, $"{AN2}-dc-*.txt", IO.SearchOption.TopDirectoryOnly)
                                                                                                                                                          For Each f In files
                                                                                                                                                              Dim disp As String = IO.Path.GetFileName(f)
                                                                                                                                                              If isLocal Then disp &= " (local)"
                                                                                                                                                              If Not displayToPath.ContainsKey(disp) Then
                                                                                                                                                                  displayToPath.Add(disp, f)
                                                                                                                                                                  options.Add(disp)
                                                                                                                                                              End If
                                                                                                                                                          Next
                                                                                                                                                      End If
                                                                                                                                                  Next

                                                                                                                                                  If options.Count = 0 Then
                                                                                                                                                      SLib.ShowCustomMessageBox($"No DocCheck script files ({AN2}-dc-*.txt) found in the configured paths.")
                                                                                                                                                      Exit Sub
                                                                                                                                                  End If

                                                                                                                                                  ' Let user pick one
                                                                                                                                                  Dim sel As String = SLib.ShowSelectionForm("Select a DocCheck script to view or edit:", $"{AN} DocCheck Scripts", options)
                                                                                                                                                  If String.IsNullOrWhiteSpace(sel) Then Exit Sub

                                                                                                                                                  Dim chosenPath As String = Nothing
                                                                                                                                                  If displayToPath.TryGetValue(sel, chosenPath) AndAlso Not String.IsNullOrWhiteSpace(chosenPath) Then
                                                                                                                                                      SLib.ShowTextFileEditor(chosenPath, $"{AN} DocCheck Script '{chosenPath}':", True, _context)
                                                                                                                                                  End If

                                                                                                                                              Catch ex As Exception
                                                                                                                                                  SLib.ShowCustomMessageBox("Error while listing DocCheck scripts:" & vbCrLf & ex.Message)
                                                                                                                                              End Try
                                                                                                                                          End Sub)
                If answer = 1 Then


                    If INI_AllowLegacyDocFiles Then
                        DragDropFormLabel = "Document files (.txt, .doc, .docx, .xlsx, .pdf), Powerpoint (.pptx), email (.msg, .eml)."
                        DragDropFormFilter =
                                "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                "Rich Text Files (*.rtf)|*.rtf|" &
                                "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                "PDF Files (*.pdf)|*.pdf|" &
                                "PowerPoint Files (*.pptx)|*.pptx|" &
                                "Email Files (*.msg;*.eml)|*.msg;*.eml"
                    Else
                        DragDropFormLabel = "Document files (.txt, .docx, .xlsx, .pdf), Powerpoint (.pptx), email (.msg, .eml)."
                        DragDropFormFilter =
                                "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                "Rich Text Files (*.rtf)|*.rtf|" &
                                "Word Documents (*.docx)|*.docx|" &
                                "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                "PDF Files (*.pdf)|*.pdf|" &
                                "PowerPoint Files (*.pptx)|*.pptx|" &
                                "Email Files (*.msg;*.eml)|*.msg;*.eml"
                    End If

                    Dim FilePath As String = GetFileName()
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    If String.IsNullOrWhiteSpace(FilePath) Then
                        ShowCustomMessageBox("No file has been selected - will abort.")
                        Return
                    End If

                    Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()
                    FromFile = Await GetFileContent(FilePath, False, True, True)

                    If FromFile.StartsWith("Error:") Then
                        ShowCustomMessageBox(FromFile)
                        Return
                    End If
                    If String.IsNullOrWhiteSpace(FromFile) Then
                        ShowCustomMessageBox("The file you provided did not contain any text - will abort.")
                        Return
                    End If
                ElseIf answer = 0 Then
                    Return
                End If
            End If



            Dim textToAnalyze As System.String = Nothing

            If FromFile <> "" Then
                textToAnalyze = FromFile
            ElseIf currentSelection.Range IsNot Nothing AndAlso currentSelection.Range.Text IsNot Nothing Then
                Dim selectedText As System.String = currentSelection.Range.Text
                If selectedText IsNot Nothing AndAlso selectedText.Trim().Length > 0 Then
                    textToAnalyze = selectedText
                End If
            End If

            If textToAnalyze Is Nothing Then
                Dim fullText As System.String = doc.Content.Text
                If fullText Is Nothing OrElse fullText.Trim().Length = 0 Then
                    ShowCustomMessageBox("There is no text to analyze.")
                    Return
                End If
                If ShowCustomYesNoBox("Select entire document for analysis?", "Yes", "No, abort") = 1 Then
                    app.Selection.WholeStory() ' this updates app.Selection in-place
                    currentSelection = app.Selection                         ' keep the Selection object
                    textToAnalyze = currentSelection.Range.Text              ' and use its text
                Else
                    Return
                End If
            End If

            ' 2) Load all RuleSets from provided paths
            Dim allRuleSets As System.Collections.Generic.List(Of DocCheckRuleSet) = LoadRuleSets(DocCheckPath, DocCheckPathLocal)
            If allRuleSets Is Nothing OrElse allRuleSets.Count = 0 Then
                ShowCustomMessageBox($"No DocCheck rule sets were found. Place files named '{AN2}-dc-*.txt' into your configured path(s).")
                Return
            End If

            ' Build dropdown options (display) and map to objects
            allRuleSets.Sort(Function(a As DocCheckRuleSet, b As DocCheckRuleSet) System.String.Compare(a.Title, b.Title, System.StringComparison.OrdinalIgnoreCase))
            Dim displayToSet As System.Collections.Generic.Dictionary(Of System.String, DocCheckRuleSet) =
            New System.Collections.Generic.Dictionary(Of System.String, DocCheckRuleSet)(System.StringComparer.OrdinalIgnoreCase)
            Dim displayOptions As System.Collections.Generic.List(Of System.String) = New System.Collections.Generic.List(Of System.String)()
            For Each rs As DocCheckRuleSet In allRuleSets
                Dim display As System.String = rs.Title
                If rs.IsLocal = True Then
                    display &= " (local)"
                End If
                ' ensure uniqueness of display text
                Dim uniqueDisplay As System.String = MakeUniqueDisplay(display, displayToSet.Keys)
                displayToSet(uniqueDisplay) = rs
                displayOptions.Add(uniqueDisplay)
            Next

            ' 3) Collect parameters in ONE form
            Dim defaultRuleSetDisplay As System.String = If(displayOptions.Count > 0, displayOptions(0), System.String.Empty)
            Dim checkOnlyOneClause As System.Boolean = False
            Dim addAdditionalContext As System.Boolean = False
            Dim doBubbles As Boolean? = True
            Dim doSummary As Boolean? = True

            If FromFile <> "" Then
                doBubbles = CType(Nothing, Boolean?)
                doSummary = CType(Nothing, Boolean?)
                do2ndModel = CType(Nothing, Boolean?)
            End If

            OtherPrompt = ""
            OutputLanguage = INI_Language1

            Dim p0 As SLib.InputParameter = New SLib.InputParameter("Rule Set to use (playbook)", defaultRuleSetDisplay)
            p0.Options = New System.Collections.Generic.List(Of System.String)(displayOptions)
            Dim p1 As SLib.InputParameter = New SLib.InputParameter("Check as only one clause", checkOnlyOneClause)
            Dim p2 As SLib.InputParameter = New SLib.InputParameter("Add additional context", addAdditionalContext)
            Dim p3 As SLib.InputParameter = New SLib.InputParameter("Output as Word bubbles", doBubbles)
            Dim p4 As SLib.InputParameter = New SLib.InputParameter("Bubbles multiclause summary", doSummary)

            Dim p5 As SLib.InputParameter
            If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                do2ndModel = CBool(False)
                p5 = New SLib.InputParameter("Use a secondary model", do2ndModel)
            ElseIf INI_SecondAPI Then
                do2ndModel = CBool(False)
                p5 = New SLib.InputParameter("Use the secondary model", do2ndModel)
            Else
                p5 = New SLib.InputParameter("Use the secondary model", do2ndModel)
            End If

            Dim p6 As SLib.InputParameter = New SLib.InputParameter("Other instructions", OtherPrompt)
            Dim p7 As SLib.InputParameter = New SLib.InputParameter("Language of output", OutputLanguage)

            Dim params() As SLib.InputParameter = {p0, p1, p2, p3, p4, p5, p6, p7}

            If ShowCustomVariableInputForm("Please set the DocCheck parameters:", AN & " DocCheck", params) = False Then
                Return
            End If

            ' Read back values
            Dim chosenDisplay As System.String = System.Convert.ToString(params(0).Value)
            checkOnlyOneClause = System.Convert.ToBoolean(params(1).Value)
            addAdditionalContext = System.Convert.ToBoolean(params(2).Value)

            Dim newBubbles = params(3).Value
            If TypeOf newBubbles Is Boolean Then
                doBubbles = CBool(newBubbles)
            Else
                doBubbles = CBool(False)
            End If
            Dim newSummary = params(4).Value
            If TypeOf (newSummary) Is Boolean Then
                doSummary = CBool(newSummary)
            Else
                doSummary = CBool(False)
            End If

            Dim SecondModel = params(5).Value
            If TypeOf (SecondModel) Is Boolean Then
                do2ndModel = CBool(SecondModel)
            Else
                do2ndModel = CBool(False)
            End If
            OtherPrompt = System.Convert.ToString(params(6).Value)
            OutputLanguage = System.Convert.ToString(params(7).Value)

            ' Resolve selected rule set
            Dim chosenRuleSet As DocCheckRuleSet = Nothing
            If chosenDisplay IsNot Nothing AndAlso displayToSet.TryGetValue(chosenDisplay, chosenRuleSet) = False Then
                ShowCustomMessageBox("Selected rule set could not be resolved (check file with rules) - will abort.")
                Return
            End If
            If chosenRuleSet Is Nothing Then
                ShowCustomMessageBox("No rule set was selected - will abort.")
                Return
            End If

            ' 4) Only now gather additional context if requested
            Dim insertDocs As System.String = ""
            If addAdditionalContext = True Then
                insertDocs = GatherSelectedDocuments(True, False)
                If insertDocs Is Nothing OrElse insertDocs.Trim().Length = 0 Then
                    ShowCustomMessageBox("No content was found or an error occurred in gathering the additional document(s) - will abort.")
                    Return
                ElseIf insertDocs.StartsWith("ERROR", System.StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox("An error occured gathering the additional document(s) (" & insertDocs.Substring(6).Trim() & ") - will abort.")
                    Return
                ElseIf insertDocs.StartsWith("NONE", System.StringComparison.OrdinalIgnoreCase) Then
                    ShowCustomMessageBox("There are no other documents to add - will abort.")
                    Return
                End If
            End If

            ' 5) Re-sync to the *current* selection right before running (user might have changed it)
            currentSelection = app.Selection
            If FromFile = "" AndAlso (currentSelection.Range Is Nothing OrElse currentSelection.Range.Text Is Nothing OrElse currentSelection.Range.Text.Trim().Length = 0) Then
                ShowCustomMessageBox("Word selection is not available.")
                Return
            End If

            If FromFile = "" Then textToAnalyze = currentSelection.Range.Text

            If do2ndModel Then

                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The secondary model could not be loaded - aborting.")
                        Return
                    End If
                End If

            End If

            ' 6) Run (pass the Selection as required by your SetBubbles pipeline)
            If checkOnlyOneClause Then
                Dim Response As String = Await RunIsolatedClause(chosenRuleSet, textToAnalyze, insertDocs, currentSelection, doBubbles, do2ndModel)
            Else
                Dim Response As String = Await RunSetOfClauses(chosenRuleSet, textToAnalyze, insertDocs, currentSelection, doBubbles, do2ndModel, doSummary, DocCheckPath, DocCheckPathLocal)
            End If


        Catch ex As System.Exception
            ShowCustomMessageBox("Error in DocCheck procedure: " & ex.Message)
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ' ========================= RuleSet Loading =========================
    ''' <summary>
    ''' Loads DocCheck rule sets from the provided global and local directories.
    ''' </summary>
    ''' <param name="pathGlobal">Expanded global DocCheck path.</param>
    ''' <param name="pathLocal">Expanded local DocCheck path.</param>
    ''' <returns>Combined list of parsed DocCheckRuleSet instances.</returns>
    Private Function LoadRuleSets(ByVal pathGlobal As System.String, ByVal pathLocal As System.String) As System.Collections.Generic.List(Of DocCheckRuleSet)
        Dim list As System.Collections.Generic.List(Of DocCheckRuleSet) = New System.Collections.Generic.List(Of DocCheckRuleSet)()

        Dim candidates As System.Collections.Generic.List(Of System.Tuple(Of System.String, System.Boolean)) =
        New System.Collections.Generic.List(Of System.Tuple(Of System.String, System.Boolean))()

        If pathGlobal IsNot Nothing AndAlso pathGlobal.Trim().Length > 0 AndAlso System.IO.Directory.Exists(pathGlobal) = True Then
            For Each f As System.String In EnumerateDocCheckFiles(pathGlobal)
                candidates.Add(New System.Tuple(Of System.String, System.Boolean)(f, False))
            Next
        End If

        If pathLocal IsNot Nothing AndAlso pathLocal.Trim().Length > 0 AndAlso System.IO.Directory.Exists(pathLocal) = True Then
            For Each f As System.String In EnumerateDocCheckFiles(pathLocal)
                candidates.Add(New System.Tuple(Of System.String, System.Boolean)(f, True))
            Next
        End If

        For Each t As System.Tuple(Of System.String, System.Boolean) In candidates
            list.AddRange(ParseDocCheckFile(t.Item1, t.Item2))
        Next

        Return list
    End Function

    ''' <summary>
    ''' Enumerates DocCheck script files within the given folder.
    ''' </summary>
    ''' <param name="folder">Folder to scan for AN2-dc-*.txt files.</param>
    ''' <returns>File paths matching the DocCheck naming pattern.</returns>
    Private Function EnumerateDocCheckFiles(ByVal folder As System.String) As System.Collections.Generic.IEnumerable(Of System.String)
        Dim matches As System.Collections.Generic.List(Of System.String) = New System.Collections.Generic.List(Of System.String)()
        Try
            For Each f As System.String In System.IO.Directory.EnumerateFiles(folder, $"{AN2}-dc-*.txt", System.IO.SearchOption.TopDirectoryOnly)
                matches.Add(f)
            Next
        Catch ex As System.Exception
            ShowCustomMessageBox("Failed to enumerate files in '" & folder & "': " & ex.Message)
        End Try
        Return matches
    End Function

    ' ==== 2) ParseDocCheckFile: add file/segment-level variables & logic ====
    ''' <summary>
    ''' Parses a DocCheck rules file and yields rule sets for each segment declared in the file.
    ''' </summary>
    ''' <param name="filePath">Full path to the *.txt rule script.</param>
    ''' <param name="isLocal">Indicates whether the source file came from the local override directory.</param>
    ''' <returns>List of DocCheckRuleSet instances extracted from the file.</returns>
    Private Function ParseDocCheckFile(ByVal filePath As System.String, ByVal isLocal As System.Boolean) As System.Collections.Generic.List(Of DocCheckRuleSet)
        Dim sets As New System.Collections.Generic.List(Of DocCheckRuleSet)()

        Try
            ' File-level defaults
            Dim fileDefaultClausePrompt As System.String = Nothing
            Dim fileDefaultMultiPrompt As System.String = Nothing
            Dim fileDefaultNotice As System.String = Nothing
            Dim fileDefaultSummaryPrompt As System.String = Nothing
            Dim fileDefaultSummaryPromptBubbles As System.String = Nothing
            Dim fileDefaultMarkdownBubbles As System.Nullable(Of System.Boolean) = Nothing

            ' Segment state
            Dim currentTitle As System.String = Nothing
            Dim jsonBuilder As New System.Text.StringBuilder()
            Dim segClausePrompt As System.String = Nothing
            Dim segMultiPrompt As System.String = Nothing
            Dim segNotice As System.String = Nothing
            Dim segSummaryPrompt As System.String = Nothing
            Dim segSummaryPromptBubbles As System.String = Nothing
            Dim segMarkdownBubbles As System.Nullable(Of System.Boolean) = Nothing

            Dim FlushCurrent As System.Action =
        Sub()
            Dim raw As System.String = jsonBuilder.ToString().Trim()
            If currentTitle IsNot Nothing AndAlso raw.Length > 0 Then
                Dim effClause As System.String = If(segClausePrompt, fileDefaultClausePrompt)
                Dim effMulti As System.String = If(segMultiPrompt, fileDefaultMultiPrompt)
                Dim effNotice As System.String = If(segNotice, fileDefaultNotice)
                Dim effSummary As System.String = If(segSummaryPrompt, fileDefaultSummaryPrompt)
                Dim effSummaryB As System.String = If(segSummaryPromptBubbles, fileDefaultSummaryPromptBubbles)
                Dim effMarkdownBubbles As System.Nullable(Of System.Boolean) = If(segMarkdownBubbles.HasValue, segMarkdownBubbles, fileDefaultMarkdownBubbles) ' << NEW

                sets.Add(CreateRuleSet(currentTitle,
                                        raw,
                                        filePath,
                                        isLocal,
                                        effClause,
                                        effMulti,
                                        effNotice,
                                        effSummary,
                                        effSummaryB,
                                        effMarkdownBubbles))
            End If
            jsonBuilder.Clear()
            segClausePrompt = Nothing
            segMultiPrompt = Nothing
            segNotice = Nothing
            segSummaryPrompt = Nothing
            segSummaryPromptBubbles = Nothing
            segMarkdownBubbles = Nothing
        End Sub

            For Each rawLine As System.String In System.IO.File.ReadLines(filePath)
                If rawLine Is Nothing Then Continue For
                Dim line As System.String = rawLine.Trim()
                If line.StartsWith(";", System.StringComparison.Ordinal) Then Continue For

                ' Notice
                Dim noticeValue As System.String = Nothing
                If TryParseNoticeLine(line, noticeValue) Then
                    If currentTitle IsNot Nothing Then
                        segNotice = noticeValue
                    Else
                        fileDefaultNotice = noticeValue
                    End If
                    Continue For
                End If

                ' Prompts
                Dim k As System.String = Nothing
                Dim v As System.String = Nothing
                If TryParsePromptLine(line, k, v) Then
                    If currentTitle IsNot Nothing Then
                        If k.Equals("SP_DocCheck_Clause", StringComparison.OrdinalIgnoreCase) Then
                            segClausePrompt = v
                        ElseIf k.Equals("SP_Docheck_MultiClause", StringComparison.OrdinalIgnoreCase) Then
                            segMultiPrompt = v
                        ElseIf k.Equals("SP_DocCheck_MultiClauseSum", StringComparison.OrdinalIgnoreCase) Then
                            segSummaryPrompt = v
                        ElseIf k.Equals("SP_DocCheck_MultiClauseSum_Bubbles", StringComparison.OrdinalIgnoreCase) Then
                            segSummaryPromptBubbles = v
                        End If
                    Else
                        If k.Equals("SP_DocCheck_Clause", StringComparison.OrdinalIgnoreCase) Then
                            fileDefaultClausePrompt = v
                        ElseIf k.Equals("SP_DocCheck_MultiClause", StringComparison.OrdinalIgnoreCase) Then
                            fileDefaultMultiPrompt = v
                        ElseIf k.Equals("SP_DocCheck_MultiClauseSum", StringComparison.OrdinalIgnoreCase) Then
                            fileDefaultSummaryPrompt = v
                        ElseIf k.Equals("SP_DocCheck_MultiClauseSum_Bubbles", StringComparison.OrdinalIgnoreCase) Then
                            fileDefaultSummaryPromptBubbles = v
                        End If
                    End If
                    Continue For
                End If

                ' MarkdownBubbles switch
                Dim mdVal As System.Nullable(Of System.Boolean) = Nothing
                If TryParseMarkdownBubblesLine(line, mdVal) Then
                    If currentTitle IsNot Nothing Then
                        segMarkdownBubbles = mdVal
                    Else
                        fileDefaultMarkdownBubbles = mdVal
                    End If
                    Continue For
                End If

                ' Start of segment
                If line.StartsWith("[", StringComparison.Ordinal) AndAlso line.EndsWith("]", StringComparison.Ordinal) AndAlso Not line.Contains("{") Then
                    FlushCurrent()
                    currentTitle = line.Substring(1, line.Length - 2).Trim()
                    Continue For
                End If

                If currentTitle IsNot Nothing Then
                    jsonBuilder.AppendLine(rawLine)
                End If
            Next

            FlushCurrent()

        Catch ex As System.Exception
            ShowCustomMessageBox("Failed to parse '" & filePath & "': " & ex.Message)
        End Try

        Return sets
    End Function

    ''' <summary>
    ''' Parses a MarkdownBubbles directive and maps the value to a nullable Boolean.
    ''' </summary>
    ''' <param name="line">Line of text possibly containing the MarkdownBubbles setting.</param>
    ''' <param name="valueOut">Receives True/False when parsing succeeds.</param>
    ''' <returns>True when the switch is recognized; otherwise False.</returns>
    Private Function TryParseMarkdownBubblesLine(ByVal line As System.String, ByRef valueOut As System.Nullable(Of System.Boolean)) As System.Boolean
        valueOut = Nothing
        If line Is Nothing Then Return False

        Dim m As System.Text.RegularExpressions.Match =
        System.Text.RegularExpressions.Regex.Match(
            line,
            "^\s*MarkdownBubbles\s*=\s*(.+)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        )

        If m Is Nothing OrElse Not m.Success Then Return False

        Dim raw As String = m.Groups(1).Value.Trim()
        If String.IsNullOrWhiteSpace(raw) Then Return False

        ' Take first token (stops at whitespace or ; or # for trailing comments)
        Dim tokMatch = System.Text.RegularExpressions.Regex.Match(raw.ToLowerInvariant(), "^\s*([^\s;#]+)")
        If Not tokMatch.Success Then Return False

        Dim tok As String = tokMatch.Groups(1).Value

        Select Case tok
            Case "true", "yes", "ja", "1", "on"
                valueOut = True : Return True
            Case "false", "no", "nein", "0", "off"
                valueOut = False : Return True
            Case Else
                Return False
        End Select
    End Function

    ''' <summary>
    ''' Parses a Notice directive and extracts its string value.
    ''' </summary>
    ''' <param name="line">Line that might contain a Notice assignment.</param>
    ''' <param name="noticeOut">Receives the trimmed notice text.</param>
    ''' <returns>True when a notice is found, False otherwise.</returns>
    Private Function TryParseNoticeLine(ByVal line As System.String, ByRef noticeOut As System.String) As System.Boolean
        noticeOut = Nothing
        If line Is Nothing Then
            Return False
        End If

        Dim m As System.Text.RegularExpressions.Match =
        System.Text.RegularExpressions.Regex.Match(
            line,
            "^\s*Notice\s*=\s*(.*)$",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        )

        If m IsNot Nothing AndAlso m.Success = True Then
            noticeOut = m.Groups(1).Value.Trim()
            Return True
        End If

        Return False
    End Function

    ''' <summary>
    ''' Creates a DocCheckRuleSet populated with prompts, notice text, and Markdown overrides.
    ''' </summary>
    ''' <param name="title">Display title of the rule set segment.</param>
    ''' <param name="rawJson">Raw JSON payload provided to the LLM.</param>
    ''' <param name="sourcePath">Path of the originating script file.</param>
    ''' <param name="isLocal">Marks the rule as coming from the local directory.</param>
    ''' <param name="fileOrSegmentClausePrompt">Effective clause prompt override.</param>
    ''' <param name="fileOrSegmentMultiPrompt">Effective multi-clause prompt override.</param>
    ''' <param name="fileOrSegmentNotice">Effective notice message.</param>
    ''' <param name="fileOrSegmentSummaryPrompt">Effective textual summary prompt.</param>
    ''' <param name="fileOrSegmentSummaryPromptBubbles">Effective bubble summary prompt.</param>
    ''' <param name="fileOrSegmentMarkdownBubbles">Markdown override flag.</param>
    ''' <returns>An initialized DocCheckRuleSet.</returns>
    Private Function CreateRuleSet(ByVal title As System.String,
                               ByVal rawJson As System.String,
                               ByVal sourcePath As System.String,
                               ByVal isLocal As System.Boolean,
                               ByVal fileOrSegmentClausePrompt As System.String,
                               ByVal fileOrSegmentMultiPrompt As System.String,
                               ByVal fileOrSegmentNotice As System.String,
                               ByVal fileOrSegmentSummaryPrompt As System.String,
                               ByVal fileOrSegmentSummaryPromptBubbles As System.String,
                               ByVal fileOrSegmentMarkdownBubbles As System.Nullable(Of System.Boolean)) As DocCheckRuleSet  ' << CHANGED

        Dim rs As New DocCheckRuleSet()
        rs.Title = title
        rs.SourcePath = sourcePath
        rs.IsLocal = isLocal
        rs.RawJson = rawJson
        rs.RecordJsons = ExtractRecordJsonStrings(rawJson)

        rs.ClausePrompt = If(Not System.String.IsNullOrWhiteSpace(fileOrSegmentClausePrompt), fileOrSegmentClausePrompt, SP_DocCheck_Clause)
        rs.MultiClausePrompt = If(Not System.String.IsNullOrWhiteSpace(fileOrSegmentMultiPrompt), fileOrSegmentMultiPrompt, SP_DocCheck_MultiClause)
        rs.NoticeText = If(fileOrSegmentNotice, Nothing)

        rs.SummaryPrompt = If(Not System.String.IsNullOrWhiteSpace(fileOrSegmentSummaryPrompt), fileOrSegmentSummaryPrompt, SP_DocCheck_MultiClauseSum)
        rs.SummaryPrompt_Bubbles = If(Not System.String.IsNullOrWhiteSpace(fileOrSegmentSummaryPromptBubbles), fileOrSegmentSummaryPromptBubbles, SP_DocCheck_MultiClauseSum_Bubbles)

        rs.MarkdownBubblesOverride = fileOrSegmentMarkdownBubbles

        Return rs
    End Function

    ''' <summary>
    ''' Parses prompt override lines and extracts the key/value pair for DocCheck prompt settings.
    ''' </summary>
    ''' <param name="line">Line that might contain an SP_DocCheck assignment.</param>
    ''' <param name="keyOut">Receives the matched key.</param>
    ''' <param name="valueOut">Receives the trimmed value.</param>
    ''' <returns>True if the pattern matches, False otherwise.</returns>
    Private Function TryParsePromptLine(ByVal line As System.String,
                                    ByRef keyOut As System.String,
                                    ByRef valueOut As System.String) As System.Boolean
        keyOut = Nothing
        valueOut = Nothing
        If line Is Nothing Then Return False

        Dim m As System.Text.RegularExpressions.Match =
        System.Text.RegularExpressions.Regex.Match(
            line,
            "^\s*(SP_DocCheck_(Clause|MultiClause|MultiClauseSum|MultiClauseSum_Bubbles))\s*=\s*(.*)$",   ' << MODIFIED
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        )

        If m IsNot Nothing AndAlso m.Success = True Then
            keyOut = m.Groups(1).Value
            valueOut = m.Groups(3).Value.Trim()
            Return True
        End If
        Return False
    End Function

    ' ========================= Execution =========================
    ''' <summary>
    ''' Executes an isolated clause run using the selected rule set and returns the LLM response.
    ''' </summary>
    ''' <param name="ruleSet">Rule set providing prompts and notices.</param>
    ''' <param name="textToAnalyze">Text being analyzed.</param>
    ''' <param name="insertDocs">Additional context appended to the user prompt.</param>
    ''' <param name="Selection">Active Word selection, required for bubble insertion.</param>
    ''' <param name="PutinBubbles">True to insert comment bubbles; False to show a window.</param>
    ''' <param name="UseSecondAPI">True to route the request through the secondary model/API.</param>
    ''' <returns>LLM response text (possibly empty when aborted).</returns>
    Private Async Function RunIsolatedClause(ByVal ruleSet As DocCheckRuleSet,
                                         ByVal textToAnalyze As System.String,
                                         ByVal insertDocs As System.String,
                                         ByVal Selection As Microsoft.Office.Interop.Word.Selection,
                                         ByVal PutinBubbles As System.Boolean,
                                         ByVal UseSecondAPI As System.Boolean) As System.Threading.Tasks.Task(Of System.String)
        Try

            If ruleSet.ClausePrompt.Trim().ToUpper = "X" Then
                ShowCustomMessageBox("The selected rule set does not work with the option 'Check as only one clause' - aborting.")
                Return System.String.Empty
            End If

            Dim systemPrompt As System.String =
            ruleSet.ClausePrompt & System.Environment.NewLine &
            (If(PutinBubbles, " " & SP_Add_Bubbles & System.Environment.NewLine, "")) &
            "<RULESET>" & ruleSet.RawJson & "</RULESET>"
            Dim userPrompt As System.String = "<TEXTTOANALYZE>" & textToAnalyze & "</TEXTTOANALYZE> "
            If Not System.String.IsNullOrWhiteSpace(insertDocs) Then
                userPrompt &= System.Environment.NewLine & "FURTHER CONTEXT: " & System.Environment.NewLine & insertDocs
            End If

            Dim MarkDown As Boolean = INI_MarkdownBubbles
            If ruleSet.MarkdownBubblesOverride.HasValue Then MarkDown = ruleSet.MarkdownBubblesOverride.Value
            If MarkDown Then FormatInstruction = SP_Add_Bubbles_Format Else FormatInstruction = ""

            Dim answer As System.String = Await LLM(InterpolateAtRuntime(systemPrompt), userPrompt, "", "", 0, UseSecondAPI)

            answer = answer.Trim()

            If Len(answer) > 3 Then
                If PutinBubbles Then
                    ' Temporarily override INI_MarkdownBubbles if requested
                    Dim oldMd As Boolean = INI_MarkdownBubbles
                    Dim changed As Boolean = False
                    If ruleSet.MarkdownBubblesOverride.HasValue Then
                        INI_MarkdownBubbles = ruleSet.MarkdownBubblesOverride.Value
                        changed = True
                    End If
                    Try
                        Dim docRef As Word.Document = Selection.Range.Document
                        Dim endPos As Integer = Selection.Range.End
                        SetBubbles(answer, Selection, False)
                        If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                            ShowCustomMessageBox("DocCheck analysis completed." & vbCrLf & vbCrLf & ruleSet.NoticeText)
                        End If
                        ' Add the Notice as a final bubble at the selection end (no text applied)
                        If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                            AddNoticeBubbleAt(docRef, endPos, ruleSet.NoticeText)
                        End If
                    Finally
                        If changed Then INI_MarkdownBubbles = oldMd
                    End Try
                Else
                    If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                        answer &= vbCrLf & vbCrLf & ruleSet.NoticeText
                    End If
                    ShowDocCheckResult(answer)
                End If
            End If

            Return answer

        Catch ex As System.Exception
            ShowCustomMessageBox("DocCheck 'One Clause' run failed: " & ex.Message)
            Return System.String.Empty
        End Try
    End Function

    ' Multi clause:

    ' Place inside Class ThisAddIn (near other private fields)
    Private _docCheckRunning As Boolean = False
    Private _docCheckAnchorStart As Integer = 0

    ' Re-entrant blocker compatible with SetBubbles()
    Private _suppressBalloonClicks As Integer = 0
    Private _handlingSelectionChange As Boolean = False
    Private _docCheckAnchorDoc As Word.Document
    Private _docCheckSelectionApplication As Microsoft.Office.Interop.Word.Application

    Private Sub StartDocCheckSelectionTracking()
        If Not WordSelectionTrackingEnabled Then Return
        If GetWordEventSubscriptionStatus("DocCheck.WindowSelectionChange") <> WordEventSubscriptionStatus.Inactive Then Return
        Try
            _docCheckSelectionApplication = Application
            If Not WordSelectionTrackingEnabled Then
                StopDocCheckSelectionTracking()
                Return
            End If
            If _docCheckSelectionApplication Is Nothing Then Throw New System.InvalidOperationException("The Word event source is unavailable.")
            SetWordEventSubscription("DocCheck.WindowSelectionChange", True,
                Sub()
                    AddHandler _docCheckSelectionApplication.WindowSelectionChange, AddressOf WordApp_WindowSelectionChange
                End Sub)
        Catch ex As System.Exception
            LogWordCommandDiagnostic("DocCheck.SelectionTracking.Startup", ex.ToString())
        End Try
    End Sub

    Private Sub StopDocCheckSelectionTracking()
        If _docCheckSelectionApplication Is Nothing Then Return
        If SetWordEventSubscription("DocCheck.WindowSelectionChange", False,
            Sub()
                RemoveHandler _docCheckSelectionApplication.WindowSelectionChange, AddressOf WordApp_WindowSelectionChange
            End Sub) Then
            _docCheckSelectionApplication = Nothing
        End If
    End Sub

    ''' <summary>
    ''' Redirects the Word selection back to the guarded anchor range when balloon clicks are suppressed.
    ''' </summary>
    Private Sub RedirectSelectionToAnchor()
        Try
            If _docCheckAnchorDoc Is Nothing Then Exit Sub
            Dim p As Integer = System.Math.Max(0, System.Math.Min(_docCheckAnchorDoc.Content.End - 1, _docCheckAnchorStart))
            Dim r As Word.Range = _docCheckAnchorDoc.Range(Start:=p, End:=p)
            r.Select()
        Catch
            ' ignore
        End Try
    End Sub

    ' Keep selection out of the comment balloons while a guarded run is active
    ''' <summary>
    ''' Handles Word selection changes to keep focus out of the comments pane when DocCheck protections are active.
    ''' </summary>
    ''' <param name="Sel">Current Word selection supplied by the event.</param>
    Private Sub WordApp_WindowSelectionChange(ByVal Sel As Microsoft.Office.Interop.Word.Selection)
        If Not WordSelectionTrackingEnabled OrElse GetWordEventSubscriptionStatus("DocCheck.WindowSelectionChange") <> WordEventSubscriptionStatus.Active Then Return
        If _handlingSelectionChange Then Exit Sub

        ' Activate guard if either a run is ongoing or a suppress request is active
        If _suppressBalloonClicks <= 0 AndAlso Not _docCheckRunning Then Exit Sub

        Try
            _handlingSelectionChange = True

            Dim inComments As Boolean = False
            Try
                inComments = (Sel.Range.StoryType = Word.WdStoryType.wdCommentsStory) _
                         OrElse CBool(Sel.Information(Word.WdInformation.wdInCommentPane))
            Catch
                inComments = False
            End Try

            If inComments Then
                RedirectSelectionToAnchor()
            End If
        Finally
            _handlingSelectionChange = False
        End Try
    End Sub

    ''' <summary>
    ''' Runs the multi-clause pipeline, iterating over each record JSON and optionally inserting bubbles or generating summaries.
    ''' </summary>
    ''' <param name="ruleSet">Rule set containing the record JSON collection.</param>
    ''' <param name="textToAnalyze">Full text block to evaluate.</param>
    ''' <param name="insertDocs">Additional context text gathered from other documents.</param>
    ''' <param name="Selection">Word selection guiding bubble placement.</param>
    ''' <param name="PutInBubbles">True to emit comment bubbles for each response.</param>
    ''' <param name="UseSecondAPI">True when the request should use the secondary model/API.</param>
    ''' <param name="DoSummary">True to request a summary bubble or report.</param>
    ''' <param name="DocCheckPath">Configured global DocCheck path (used for UI callbacks).</param>
    ''' <param name="DocCheckPathLocal">Configured local DocCheck path (used for UI callbacks).</param>
    ''' <returns>Concatenated LLM responses prior to optional summarization.</returns>
    Private Async Function RunSetOfClauses(ByVal ruleSet As DocCheckRuleSet,
                                       ByVal textToAnalyze As System.String,
                                       ByVal insertDocs As System.String,
                                       ByVal Selection As Microsoft.Office.Interop.Word.Selection,
                                       ByVal PutInBubbles As System.Boolean,
                                       ByVal UseSecondAPI As System.Boolean,
                                       ByVal DoSummary As System.Boolean,
                                       ByVal DocCheckPath As String, ByVal DocCheckPathLocal As String) As System.Threading.Tasks.Task(Of System.String)
        Dim sw As Stopwatch = Stopwatch.StartNew()
        Dim prevScreenUpdating As Boolean = True
        Try
            ' Guard against user clicking into comment balloons during the run
            prevScreenUpdating = wordApp.ScreenUpdating
            'wordApp.ScreenUpdating = False
            _docCheckRunning = True

            Dim records As System.Collections.Generic.List(Of System.String) = ruleSet.RecordJsons
            If records Is Nothing OrElse records.Count = 0 Then
                records = ExtractRecordJsonStrings(ruleSet.RawJson)
                If records Is Nothing OrElse records.Count = 0 Then
                    Dim recordsDebugInfo As String = If(ruleSet.RecordJsons Is Nothing,
                                            "NULL",
                                            $"Count={ruleSet.RecordJsons.Count}, Content=[{String.Join("; ", ruleSet.RecordJsons)}]")
                    ShowCustomMessageBox($"This rule set contains no valid records. Expected a JSON object with 'Records', or an array/object representing records.{vbCrLf}{vbCrLf}Debug Info - RecordJsons: {recordsDebugInfo}",
                                         extraButtonText:="Edit Script Files", extraButtonAction:=Sub()
                                                                                                      Try
                                                                                                          ' Collect DocCheck script files from both paths
                                                                                                          Dim displayToPath As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                                                                                                          Dim options As New List(Of String)

                                                                                                          Dim dcPaths As New List(Of (p As String, isLocal As Boolean))
                                                                                                          If Not String.IsNullOrWhiteSpace(DocCheckPath) Then
                                                                                                              dcPaths.Add((DocCheckPath, False))
                                                                                                          End If
                                                                                                          If Not String.IsNullOrWhiteSpace(DocCheckPathLocal) Then
                                                                                                              dcPaths.Add((DocCheckPathLocal, True))
                                                                                                          End If

                                                                                                          For Each tuple In dcPaths
                                                                                                              Dim basePath = tuple.p
                                                                                                              Dim isLocal = tuple.isLocal
                                                                                                              If IO.Directory.Exists(basePath) Then
                                                                                                                  Dim files = IO.Directory.GetFiles(basePath, $"{AN2}-dc-*.txt", IO.SearchOption.TopDirectoryOnly)
                                                                                                                  For Each f In files
                                                                                                                      Dim disp As String = IO.Path.GetFileName(f)
                                                                                                                      If isLocal Then disp &= " (local)"
                                                                                                                      If Not displayToPath.ContainsKey(disp) Then
                                                                                                                          displayToPath.Add(disp, f)
                                                                                                                          options.Add(disp)
                                                                                                                      End If
                                                                                                                  Next
                                                                                                              End If
                                                                                                          Next

                                                                                                          If options.Count = 0 Then
                                                                                                              SLib.ShowCustomMessageBox($"No DocCheck script files ({AN2}-dc-*.txt) found in the configured paths.")
                                                                                                              Exit Sub
                                                                                                          End If

                                                                                                          ' Let user pick one
                                                                                                          Dim sel As String = SLib.ShowSelectionForm("Select a DocCheck script to view or edit:", $"{AN} DocCheck Scripts", options)
                                                                                                          If String.IsNullOrWhiteSpace(sel) Then Exit Sub

                                                                                                          Dim chosenPath As String = Nothing
                                                                                                          If displayToPath.TryGetValue(sel, chosenPath) AndAlso Not String.IsNullOrWhiteSpace(chosenPath) Then
                                                                                                              SLib.ShowTextFileEditor(chosenPath, $"{AN} DocCheck Script '{chosenPath}':", True, _context)
                                                                                                          End If

                                                                                                      Catch ex As Exception
                                                                                                          SLib.ShowCustomMessageBox("Error while listing DocCheck scripts:" & vbCrLf & ex.Message)
                                                                                                      End Try
                                                                                                  End Sub)


                    Return System.String.Empty
                End If
            End If

            ShowProgressBarInSeparateThread(AN & " DocCheck", "Analyzing text...")
            ProgressBarModule.CancelOperation = False

            GlobalProgressMax = records.Count
            GlobalProgressValue = 0
            GlobalProgressLabel = "Analyzing rule 0 of " & records.Count

            Dim OverallAnswer As System.String = System.String.Empty
            Dim idx As System.Int32 = 0

            Dim docRef As Word.Document = Selection.Range.Document
            Dim startPos As Integer = Selection.Range.Start
            Dim endPos As Integer = Selection.Range.End

            ' Anchor for selection bounce-back and re-entrant suppression
            _docCheckAnchorDoc = docRef
            _suppressBalloonClicks += 1

            ' Snapshot anchor used by selection-change guard
            _docCheckAnchorStart = startPos

            ' Apply MarkdownBubbles override only for bubble workflows
            Dim oldMd As Boolean = INI_MarkdownBubbles
            Dim changed As Boolean = False
            If PutInBubbles AndAlso ruleSet.MarkdownBubblesOverride.HasValue Then
                INI_MarkdownBubbles = ruleSet.MarkdownBubblesOverride.Value
                changed = True
            End If

            Try
                For Each recordJson As System.String In records
                    If ProgressBarModule.CancelOperation = True Then
                        Dim elapsedFmt = sw.Elapsed.ToString("hh\:mm\:ss\.fff")
                        ShowCustomMessageBox("Analysis aborted by user. Elapsed: " & elapsedFmt)
                        Exit For
                    End If

                    GlobalProgressValue = idx
                    GlobalProgressLabel = "Analyzing rule " & (idx + 1).ToString() & " of " & records.Count.ToString()

                    Dim systemPrompt As System.String =
                    ruleSet.MultiClausePrompt & System.Environment.NewLine &
                    (If(PutInBubbles, " " & SP_Add_Bubbles & System.Environment.NewLine, "")) &
                    "<RULESET>" & recordJson & "</RULESET>"

                    Dim userPrompt As System.String = "<TEXTTOANALYZE>" & textToAnalyze & "</TEXTTOANALYZE> "
                    If Not System.String.IsNullOrWhiteSpace(insertDocs) Then
                        userPrompt &= System.Environment.NewLine & "FURTHER CONTEXT: " & System.Environment.NewLine & insertDocs
                    End If

                    Dim MarkDown As Boolean = INI_MarkdownBubbles
                    If ruleSet.MarkdownBubblesOverride.HasValue Then MarkDown = ruleSet.MarkdownBubblesOverride.Value
                    If MarkDown Then FormatInstruction = SP_Add_Bubbles_Format Else FormatInstruction = ""

                    Dim answer As System.String = Await LLM(InterpolateAtRuntime(systemPrompt), userPrompt, "", "", 0, UseSecondAPI)
                    answer = answer.Trim()

                    If answer.Length > 3 Then
                        If PutInBubbles Then
                            ' NOTE: SetBubbles uses Selection; the guard keeps Selection in main story to prevent COM errors if user clicks into balloons
                            SetBubbles(answer, Selection, True)
                            OverallAnswer &= answer & System.Environment.NewLine & System.Environment.NewLine
                        Else
                            OverallAnswer &= answer & System.Environment.NewLine & System.Environment.NewLine
                        End If
                    End If

                    idx += 1
                Next

                If ProgressBarModule.CancelOperation = False Then
                    If PutInBubbles = False Then
                        GlobalProgressLabel = "Creating final report..."
                        GlobalProgressValue = idx
                        Dim summaryPromptToUse As String = ruleSet.SummaryPrompt
                        Dim OverallAnswer2 As System.String = Await LLM(InterpolateAtRuntime(summaryPromptToUse), "<TEXTTOPROCESS>" & OverallAnswer & "</TEXTTOPROCESS>", "", "", 0, False)
                        ProgressBarModule.CancelOperation = True

                        Dim finalReport As System.String = If(Not System.String.IsNullOrWhiteSpace(OverallAnswer2), OverallAnswer2, OverallAnswer)
                        If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                            finalReport &= vbCrLf & vbCrLf & ruleSet.NoticeText
                        End If

                        Dim elapsedFmt = sw.Elapsed.ToString("hh\:mm\:ss\.fff")
                        finalReport &= vbCrLf & vbCrLf & "Analysis time: " & elapsedFmt

                        ShowDocCheckResult(finalReport)
                    Else
                        If DoSummary Then
                            GlobalProgressLabel = "Creating summary..."
                            GlobalProgressValue = idx
                            Dim summaryPromptBubbles As String = ruleSet.SummaryPrompt_Bubbles
                            Dim OverallAnswer2 As System.String = Await LLM(InterpolateAtRuntime(summaryPromptBubbles), "<TEXTTOPROCESS>" & OverallAnswer & "</TEXTTOPROCESS>", "", "", 0, False)
                            OverallAnswer = OverallAnswer2.Trim()
                            If OverallAnswer2 <> "" Then
                                ' Always use startPos for summary (beginning of selection)
                                Dim summaryAnchor As Integer = startPos
                                AddNoticeBubbleAt(docRef, summaryAnchor, OverallAnswer2)
                            End If
                        End If
                        If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                            AddNoticeBubbleAt(docRef, endPos, ruleSet.NoticeText)
                        End If
                        ProgressBarModule.CancelOperation = True

                        Dim elapsedFmt = sw.Elapsed.ToString("hh\:mm\:ss\.fff")
                        Dim msg As System.String = "DocCheck analysis completed - check out the comments added (if any)."
                        If Not System.String.IsNullOrWhiteSpace(ruleSet.NoticeText) Then
                            msg &= vbCrLf & vbCrLf & ruleSet.NoticeText
                        End If
                        msg &= vbCrLf & vbCrLf & "Analysis time: " & elapsedFmt

                        ShowCustomMessageBox(msg)
                    End If
                End If

            Finally
                If changed Then INI_MarkdownBubbles = oldMd
            End Try

            ProgressBarModule.CancelOperation = True
            Return OverallAnswer

        Catch ex As System.Exception
            Dim elapsedFmt As String
            Try
                elapsedFmt = sw.Elapsed.ToString("hh\:mm\:ss\.fff")
            Catch
                elapsedFmt = ""
            End Try
            ShowCustomMessageBox("DocCheck 'Multi Clause' run failed: " & ex.Message & If(elapsedFmt <> "", " (Elapsed: " & elapsedFmt & ")", ""))
            Return System.String.Empty
        Finally
            _docCheckRunning = False
            sw.Stop()
            ' Restore selection guard nesting
            If _suppressBalloonClicks > 0 Then _suppressBalloonClicks -= 1
            wordApp.ScreenUpdating = prevScreenUpdating
        End Try
    End Function

    ' Helper to add a zero-length comment at a specific position
    ''' <summary>
    ''' Adds a Markdown-enabled notice bubble at the specified document position.
    ''' </summary>
    ''' <param name="doc">Target Word document.</param>
    ''' <param name="endPos">Character index where the zero-length comment is anchored.</param>
    ''' <param name="noticeText">Notice content inserted into the bubble.</param>
    ''' <param name="Prefix">Optional prefix prepended to the notice.</param>
    Private Sub AddNoticeBubbleAt(doc As Word.Document, endPos As Integer, noticeText As String, Optional Prefix As String = "")
        If Not INI_MarkdownBubbles Then
            LegacyAddNoticeBubbleAt(doc, endPos, noticeText)
            Return
        End If
        Try
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(noticeText) Then Return

            Using BeginMarkupAuthorScope(doc.Application)
                ' Zero-length anchor at endPos
                Dim anchor As Word.Range = doc.Range(endPos, endPos)

                ' Create an empty comment so we can insert formatted content
                Dim cmt As Word.Comment = doc.Comments.Add(anchor, "")
                Dim cr As Word.Range = cmt.Range

                ' Clear and insert a formatted prefix (e.g., "RI: ")
                cr.Text = ""

                Dim ins As Word.Range = cr.Duplicate
                ins.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                InsertMarkdownToComment(ins, $"{AN5}{Prefix}: " & noticeText)
            End Using
        Catch
            ' Swallow – adding the notice must not break the main flow
        End Try
    End Sub

    ''' <summary>
    ''' Adds a legacy plaintext notice bubble anchored to the main document story.
    ''' </summary>
    ''' <param name="doc">Target Word document.</param>
    ''' <param name="endPos">Ignored when seeking, retained for signature parity.</param>
    ''' <param name="noticeText">Notice content inserted as plain text.</param>
    ''' <param name="Prefix">Optional prefix prepended to the notice.</param>
    Private Sub LegacyAddNoticeBubbleAt(doc As Word.Document, endPos As Integer, noticeText As String, Optional Prefix As String = "")
        Try
            If doc Is Nothing OrElse String.IsNullOrWhiteSpace(noticeText) Then Return

            ' Ensure we operate in the MAIN story, ignore endPos (may be from another story)
            Dim win As Word.Window = doc.ActiveWindow
            Dim prevSeek As Word.WdSeekView = win.View.SeekView
            Try
                win.View.SeekView = Word.WdSeekView.wdSeekMainDocument

                Dim mainStory As Word.Range = doc.StoryRanges(Word.WdStoryType.wdMainTextStory)
                Dim anchor As Word.Range = mainStory.Duplicate
                anchor.Collapse(Word.WdCollapseDirection.wdCollapseEnd) ' zero-length in MAIN story
                Using BeginMarkupAuthorScope(doc.Application)
                    doc.Comments.Add(anchor, $"{AN5}{Prefix}: " & noticeText)
                End Using
            Finally
                ' Restore previous seek view
                win.View.SeekView = prevSeek
            End Try
        Catch
            ' Swallow – adding the notice must not break the main flow
        End Try
    End Sub

    ''' <summary>
    ''' Shows DocCheck results through the configured UI surfaces (markdown pane, clipboard, or modal window).
    ''' </summary>
    ''' <param name="answer">LLM response assembled from the DocCheck run.</param>
    Private Sub ShowDocCheckResult(ByVal answer As System.String)
        Dim result As String = SLib.ShowCustomWindow("The DocCheck resulted in the following findings:", answer, "", AN, False, True, True, True)
        If result <> "" And result <> "Pane" Then
            If result = "Markdown" Then
                Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                Globals.ThisAddIn.Application.Selection.TypeParagraph()
                InsertTextWithMarkdown(Globals.ThisAddIn.Application.Selection, vbCrLf & result, False)
                Dim patternx As String = "\{\{(WFLD|WENT|WFNT):.*?\}\}"
                If Regex.IsMatch(result, patternx) Then
                    Dim rng As Range = wordApp.Selection.Range
                    RestoreSpecialTextElements(rng)
                    rng.Document.Fields.Update()
                End If
            Else
                SLib.PutInClipboard(result)
            End If
        ElseIf result = "Pane" Then

            SP_MergePrompt_Cached = SP_MergePrompt
            If _uiContext IsNot Nothing Then  ' Make sure we run in the UI Thread
                _uiContext.Post(Sub(s)

                                    ShowPaneAsync(
                                                            "DocCheck Findings",
                                                            answer,
                                                            "",
                                                            AN,
                                                            noRTF:=False,
                                                            insertMarkdown:=True
                                                            )

                                End Sub, Nothing)
            Else

                ShowPaneAsync(
                                                            "DocCheck Findings",
                                                            answer,
                                                            "",
                                                            AN,
                                                            noRTF:=False,
                                                            insertMarkdown:=True
                                                            )
            End If

        End If
        Return
    End Sub

    ' Record schema:
    '   {
    '     "Topic": "String",
    '     "Issue": "String",
    '     "Criteria": [
    '       {
    '         "Condition": "String",
    '         "IfTrue":    { "Consequence": "String", "Risk": 1..3 },
    '         "IfFalse": { "Consequence": "String", "Risk": 1..3 }
    '       },
    '       ...
    '     ]
    '   }
    '
    ' We do NOT interpret record fields. We only extract per-record JSON as string
    ' (compact) and hand it verbatim to the LLM. Supported RuleSet shapes:
    '   1) { "Records": [ {record}, {record}, ... ] }
    '   2) [ {record}, {record}, ... ]
    '   3) {record}

    ''' <summary>
    ''' Extracts compact JSON strings for each rule record contained in the raw JSON payload.
    ''' </summary>
    ''' <param name="rawJson">Raw JSON text copied from the rule set file.</param>
    ''' <returns>List of per-record JSON strings; empty when parsing fails.</returns>
    Private Function ExtractRecordJsonStrings(ByVal rawJson As System.String) As System.Collections.Generic.List(Of System.String)
        Dim list As New System.Collections.Generic.List(Of System.String)()
        If System.String.IsNullOrWhiteSpace(rawJson) Then Return list

        ' 1) Normalize leading/trailing junk (BOM, zero-width, whitespace)
        Dim working As System.String = rawJson
        working = working.Trim()
        ' Remove BOM and common zero-width chars at start/end
        working = System.Text.RegularExpressions.Regex.Replace(working, "^[\uFEFF\u200B\u200C\u200D\u2060]+", "")
        working = System.Text.RegularExpressions.Regex.Replace(working, "[\u2060\u200B\u200C\u200D]+$", "")

        ' 2) Try to crop to the first bracketed JSON payload to ignore any prefix/suffix
        Dim startIdx As System.Int32 = working.IndexOfAny(New System.Char() {"["c, "{"c})
        If startIdx >= 0 Then
            Dim openCh As System.Char = working(startIdx)
            Dim closeCh As System.Char = If(openCh = "["c, "]"c, "}"c)

            Dim inString As System.Boolean = False
            Dim escaped As System.Boolean = False
            Dim depth As System.Int32 = 0
            Dim endIdx As System.Int32 = -1

            For i As System.Int32 = startIdx To working.Length - 1
                Dim ch As System.Char = working(i)

                If inString Then
                    If escaped Then
                        escaped = False
                    ElseIf ch = "\"c Then
                        escaped = True
                    ElseIf ch = """"c Then
                        inString = False
                    End If
                Else
                    If ch = """"c Then
                        inString = True
                    ElseIf ch = openCh Then
                        depth += 1
                    ElseIf ch = closeCh Then
                        depth -= 1
                        If depth = 0 Then
                            endIdx = i
                            Exit For
                        End If
                    End If
                End If
            Next

            ' If we found a balanced payload, crop to it; if not, keep original from startIdx
            If endIdx >= 0 Then
                working = working.Substring(startIdx, endIdx - startIdx + 1)
            Else
                ' Keep from startIdx but don't lose the content
                working = working.Substring(startIdx)
            End If
        End If

        ' 3) Parse with lenient bracket repair for common mistakes
        Dim token As Newtonsoft.Json.Linq.JToken = Nothing

        ' First attempt direct parse
        If Not TryParseJsonToken(working, token) Then
            Dim t As System.String = working.Trim()
            Dim repaired As System.Boolean = False

            ' Try various repair strategies
            If t.StartsWith("[", System.StringComparison.Ordinal) Then
                If t.EndsWith("}", System.StringComparison.Ordinal) Then
                    ' Array started but ends with object bracket
                    t = t.Substring(0, t.Length - 1) & "]"
                    repaired = True
                ElseIf Not t.EndsWith("]", System.StringComparison.Ordinal) Then
                    ' Array started but no closing bracket
                    t &= "]"
                    repaired = True
                End If
            ElseIf t.StartsWith("{", System.StringComparison.Ordinal) Then
                If t.EndsWith("]", System.StringComparison.Ordinal) Then
                    ' Object started but ends with array bracket
                    t = t.Substring(0, t.Length - 1) & "}"
                    repaired = True
                ElseIf Not t.EndsWith("}", System.StringComparison.Ordinal) Then
                    ' Object started but no closing bracket
                    t &= "}"
                    repaired = True
                End If
            End If

            If repaired Then
                If TryParseJsonToken(t, token) Then
                    working = t
                Else
                    ' Repair failed, return empty
                    Return list
                End If
            Else
                ' No repair attempted, return empty
                Return list
            End If
        End If

        ' 4) Extract per-record JSON strings
        If TypeOf token Is Newtonsoft.Json.Linq.JObject Then
            Dim jo As Newtonsoft.Json.Linq.JObject = CType(token, Newtonsoft.Json.Linq.JObject)

            ' CASE 1: Object with "Records" (case-insensitive)
            Dim recordsProp As Newtonsoft.Json.Linq.JProperty = jo.Properties().
        FirstOrDefault(Function(p) System.String.Equals(p.Name, "Records", System.StringComparison.OrdinalIgnoreCase))

            If recordsProp IsNot Nothing AndAlso recordsProp.Value.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                For Each rec As Newtonsoft.Json.Linq.JToken In CType(recordsProp.Value, Newtonsoft.Json.Linq.JArray)
                    list.Add(rec.ToString(Newtonsoft.Json.Formatting.None))
                Next
                Return list
            End If

            ' CASE 3: single record object
            list.Add(jo.ToString(Newtonsoft.Json.Formatting.None))
            Return list
        End If

        ' CASE 2: top-level array
        If TypeOf token Is Newtonsoft.Json.Linq.JArray Then
            For Each rec As Newtonsoft.Json.Linq.JToken In CType(token, Newtonsoft.Json.Linq.JArray)
                list.Add(rec.ToString(Newtonsoft.Json.Formatting.None))
            Next
            Return list
        End If

        Return list
    End Function

    ''' <summary>
    ''' Wraps Newtonsoft parsing to safely obtain a JToken while shielding callers from exceptions.
    ''' </summary>
    ''' <param name="s">JSON string to parse.</param>
    ''' <param name="t">Receives the parsed token on success.</param>
    ''' <returns>True when parsing succeeds; otherwise False.</returns>
    Private Function TryParseJsonToken(ByVal s As System.String, ByRef t As Newtonsoft.Json.Linq.JToken) As System.Boolean
        Try
            t = Newtonsoft.Json.Linq.JToken.Parse(s)
            Return True
        Catch ex As System.Exception
            Return False
        End Try
    End Function

    ' ========================= Types =========================
    ''' <summary>
    ''' Represents a DocCheck rule set, storing prompts, notices, and extracted record JSON payloads.
    ''' </summary>
    Private Class DocCheckRuleSet
        Public Property Title As System.String
        Public Property SourcePath As System.String
        Public Property IsLocal As System.Boolean
        Public Property RawJson As System.String
        Public Property RecordJsons As System.Collections.Generic.List(Of System.String)
        Public Property ClausePrompt As System.String
        Public Property MultiClausePrompt As System.String
        Public Property NoticeText As System.String
        Public Property SummaryPrompt As System.String
        Public Property SummaryPrompt_Bubbles As System.String
        Public Property MarkdownBubblesOverride As System.Nullable(Of System.Boolean)
    End Class

    ' ========================= Small utility =========================
    ''' <summary>
    ''' Ensures display names shown to users remain unique by appending a numeric suffix as needed.
    ''' </summary>
    ''' <param name="baseText">Original display text.</param>
    ''' <param name="existing">Existing collection of display names to compare against.</param>
    ''' <returns>A unique display string.</returns>
    Private Function MakeUniqueDisplay(ByVal baseText As System.String, ByVal existing As System.Collections.Generic.ICollection(Of System.String)) As System.String
        If existing Is Nothing OrElse existing.Contains(baseText) = False Then
            Return baseText
        End If
        Dim i As System.Int32 = 2
        While existing.Contains(baseText & " [" & i & "]")
            i += 1
        End While
        Return baseText & " [" & i & "]"
    End Function

End Class