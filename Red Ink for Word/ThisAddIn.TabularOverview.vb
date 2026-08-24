' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.TabularOverview.vb
' Purpose:
'   Orchestrates fact extraction from one or multiple documents into a Word table.
'   Loads prepared instruction/schema entries, applies manual overrides,
'   resolves merge rules, optional secondary model usage, performs extraction,
'   and writes a normalized fact table with optional date clamping, sorting,
'   formatting, and summary metadata.
'
' Architecture:
'   - Instruction/Schema Library: Text files (local/global) enumerated; each line pipe-delimited:
'       Title | Instruction | SchemaSpec | MergeEnable | MergeDateCol | MergeInstruction
'   - Parameters: Builds effective instruction, schema, date columns, clamp bounds,
'       sort parameters, OCR toggle, output language, and merge intent.
'   - Merge Resolution: Based on checkbox, manual merge key column, and library metadata
'       (backwards-compatible with MergeDateCol).
'   - Schema Handling: Manual overrides; prepared; optional AI generation.
'   - Multiple files:
'       When enabled, prompts for a folder and processes all supported files.
'       A progress UI is shown. Summary rows are appended after data.
'   - Execution: Single-file or folder batch; progress via global progress variables.
'   - Result Insertion: Creates a Word table at the cursor position or in a new document,
'       with headers, rows, optional date formatting (date/datetime only), and appended
'       summary rows.
'   - Cleanup: Restores original model configuration if a secondary model was temporarily loaded.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports Newtonsoft.Json
Imports SharedLibrary.FactExtractionService

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Sentinel display value inserted as the first dropdown option to trigger the interactive library builder.
    ''' </summary>
    Private Const TabularBuilderSentinel As String = "✨ Build new library entry..."


    ''' <summary>
    ''' Main entry point for tabular overview (fact extraction into a Word table).
    ''' Collects user parameters, resolves instruction/schema, optional secondary model usage,
    ''' executes extraction (single or multi-file) and inserts results into a Word table.
    ''' </summary>
    Public Async Sub TabularOverview()

        If INILoadFail() OrElse Not IsDocumentEditable() Then Return

        Dim useSecondApi As Boolean = False
        Dim do2ndModel As Boolean = False

        Try
            Dim displayToInstruction As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim displayToSchema As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim displayToMergeEnable As New System.Collections.Generic.Dictionary(Of String, Boolean)(StringComparer.OrdinalIgnoreCase)
            Dim displayToMergeDateCol As New System.Collections.Generic.Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            Dim displayToMergeInstruction As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)

            Dim displayOptions As New System.Collections.Generic.List(Of String)()
            Dim localPath = ExpandEnvironmentVariables(INI_ExtractorPathLocal)
            Dim globalPath = ExpandEnvironmentVariables(INI_ExtractorPath)

            Dim EnumerateInstructionFiles As Func(Of String, System.Collections.Generic.IEnumerable(Of String)) =
                Function(p As String)
                    Dim result As New System.Collections.Generic.List(Of String)
                    If String.IsNullOrWhiteSpace(p) Then Return result
                    Try
                        If Directory.Exists(p) Then
                            result.AddRange(Directory.GetFiles(p, "*.txt", SearchOption.TopDirectoryOnly))
                        ElseIf File.Exists(p) Then
                            result.Add(p)
                        End If
                    Catch
                    End Try
                    Return result
                End Function

            Dim LoadFromFile As Action(Of String, Boolean) =
                Sub(file As String, isLocal As Boolean)
                    Try
                        For Each rawLine In System.IO.File.ReadAllLines(file)
                            Dim line = If(rawLine, "").Trim()
                            If line.Length = 0 OrElse line.StartsWith(";", StringComparison.Ordinal) Then Continue For
                            Dim parts = line.Split("|"c)

                            If parts.Length < 1 Then Continue For
                            Dim title = parts(0).Trim()
                            Dim instr As String = If(parts.Length >= 2, parts(1).Trim(), "")
                            Dim schemaSpec = If(parts.Length >= 3, parts(2).Trim(), "")

                            Dim libMergeEnable As Boolean = False
                            Dim libMergeDateCol As Integer = 0
                            Dim libMergeInstr As String = ""

                            If parts.Length >= 4 Then
                                Dim s = parts(3).Trim()
                                Dim b As Boolean
                                If Boolean.TryParse(s, b) Then libMergeEnable = b
                            End If
                            If parts.Length >= 5 Then
                                Dim s = parts(4).Trim()
                                Dim n As Integer
                                If Integer.TryParse(s, n) AndAlso n > 0 Then libMergeDateCol = n
                            End If
                            If parts.Length >= 6 Then
                                libMergeInstr = parts(5).Trim()
                            End If

                            If title.Length = 0 Then Continue For
                            Dim display = title & If(isLocal, " (local)", "")
                            Dim unique = TabularMakeUniqueDisplay(display, displayToInstruction.Keys)

                            displayToInstruction(unique) = instr
                            If schemaSpec.Length > 0 Then displayToSchema(unique) = schemaSpec

                            displayToMergeEnable(unique) = libMergeEnable
                            displayToMergeDateCol(unique) = libMergeDateCol
                            displayToMergeInstruction(unique) = libMergeInstr

                            displayOptions.Add(unique)
                        Next
                    Catch
                    End Try
                End Sub

            For Each f In EnumerateInstructionFiles(localPath) : LoadFromFile(f, True) : Next
            For Each f In EnumerateInstructionFiles(globalPath) : LoadFromFile(f, False) : Next

            ' Insert the builder sentinel as the first dropdown option when a local path is available
            If Not String.IsNullOrWhiteSpace(localPath) Then
                displayOptions.Insert(0, TabularBuilderSentinel)
            End If

            Dim defaultManual = ""
            Dim defaultManualSchema = ""
            Dim defaultDateCols = ""
            Dim defaultSortCol = ""
            Dim defaultSortDir = "ASC"
            Dim defaultDoOcr As Boolean = False
            Dim defaultClampFrom As String = ""
            Dim defaultClampTo As String = ""
            Dim defaultOutputLanguage As String = ""
            Dim defaultDateOutputFormat As String = ""
            Dim defaultMergeEnable As Boolean = False
            Dim defaultMergeDateColumn As Integer = 0
            Dim defaultMergeInstruction As String = ""

            Try : defaultManual = System.Convert.ToString(My.Settings.Tabular_ManualInstruction) : Catch : End Try
            Try : defaultManualSchema = System.Convert.ToString(My.Settings.Tabular_ManualSchema) : Catch : End Try
            Try : defaultDateCols = System.Convert.ToString(My.Settings.Tabular_DateColumns) : Catch : End Try
            Try : defaultSortCol = System.Convert.ToString(My.Settings.Tabular_SortColumn) : Catch : End Try
            Try : defaultSortDir = System.Convert.ToString(My.Settings.Tabular_SortDirection) : Catch : End Try
            Try : defaultClampFrom = System.Convert.ToString(My.Settings.Tabular_DateClampFrom) : Catch : End Try
            Try : defaultClampTo = System.Convert.ToString(My.Settings.Tabular_DateClampTo) : Catch : End Try
            Try : defaultOutputLanguage = System.Convert.ToString(My.Settings.Tabular_OutputLanguage) : Catch : End Try
            Try : defaultDoOcr = My.Settings.Tabular_DoOcr : Catch : End Try
            Try : defaultDateOutputFormat = System.Convert.ToString(My.Settings.Tabular_DateOutputFormat) : Catch : End Try
            Try : defaultMergeEnable = My.Settings.Tabular_MergeEnable : Catch : defaultMergeEnable = False : End Try
            Try
                Dim tmp = System.Convert.ToString(My.Settings.Tabular_MergeDateColumn)
                Dim n As Integer
                If Integer.TryParse(tmp, n) AndAlso n > 0 Then defaultMergeDateColumn = n
            Catch
            End Try
            Try : defaultMergeInstruction = System.Convert.ToString(My.Settings.Tabular_MergeInstruction) : Catch : End Try

            Dim manualInstruction = defaultManual
            Dim manualSchemaText = defaultManualSchema
            Dim dateColumnsText = defaultDateCols
            Dim sortColumnText = defaultSortCol

            Dim sortDirection As String
            Select Case (If(defaultSortDir, "").Trim().ToUpperInvariant())
                Case "ASC", "ASCENDING" : sortDirection = "ASC"
                Case "DESC", "DESCENDING" : sortDirection = "DESC"
                Case Else : sortDirection = "ASC"
            End Select

            Dim clampFrom = defaultClampFrom
            Dim clampTo = defaultClampTo
            Dim UserOutputLanguage = defaultOutputLanguage
            Dim doOcr = defaultDoOcr
            Dim dateOutputFormat = defaultDateOutputFormat
            Dim defaultPreparedDisplay = If(displayOptions.Count > 0, displayOptions(0), "<<disabled>>")

            Dim mergeRowsViaLlm As Boolean = defaultMergeEnable
            Dim mergeDateColumn As Integer = defaultMergeDateColumn
            Dim mergeInstruction As String = defaultMergeInstruction

            Dim hasSecondary = (Not String.IsNullOrWhiteSpace(INI_AlternateModelPath)) OrElse INI_SecondAPI

            Dim p0 As SLib.InputParameter = New SLib.InputParameter("Prepared instruction set", defaultPreparedDisplay)
            If displayOptions.Count > 0 Then p0.Options = New System.Collections.Generic.List(Of String)(displayOptions)
            Dim p1 As SLib.InputParameter = If(displayOptions.Count = 0,
                                               New SLib.InputParameter("Extraction instructions", manualInstruction),
                                               New SLib.InputParameter("Manual instruction (overrides)", manualInstruction))
            Dim pSchema As New SLib.InputParameter("Manual schema (semicolon Name[:type][*]; empty=auto)", manualSchemaText)
            Dim p2 As New SLib.InputParameter("Date column indices (CSV, 1-based)", dateColumnsText)
            Dim pClampFrom As New SLib.InputParameter("Only dates on or later", clampFrom)
            Dim pClampTo As New SLib.InputParameter("Only dates up and until", clampTo)
            Dim p3 As New SLib.InputParameter("Column to sort (1-based, auto=*, empty=none)", sortColumnText)
            Dim displaySortDirection = If(sortDirection = "DESC", "Descending", "Ascending")
            Dim p4 As New SLib.InputParameter("Sort direction", displaySortDirection)
            p4.Options = New System.Collections.Generic.List(Of String) From {"Ascending", "Descending"}
            ' OCR checkbox: pass Nothing if OCR is unavailable to show disabled checkbox
            Dim ocrAvailable As Boolean = SharedMethods.IsOcrAvailable(_context)
            Dim p6 As New SLib.InputParameter("Do OCR if needed (PDFs)", If(ocrAvailable, CObj(doOcr), Nothing))
            Dim p9 As New SLib.InputParameter("Output language", UserOutputLanguage)
            Dim p10 As New SLib.InputParameter("Date format (e.g., yyyy-MM-dd; empty=default)", dateOutputFormat)
            Dim pMergeEnable As New SLib.InputParameter("Permit row merging (if requested)", mergeRowsViaLlm)
            Dim pMergeDateCol As New SLib.InputParameter("Column to merge/group on (1-based)", If(mergeDateColumn <= 0, "", mergeDateColumn.ToString()))
            Dim pMergeInstruction As New SLib.InputParameter("Additional merge instructions (optional, overrides)", mergeInstruction)
            Dim pDebugDiagnostics As New SLib.InputParameter("Debug failed/no-result extractions to Desktop log", False)

            Dim p11 As SLib.InputParameter = Nothing
            If hasSecondary Then
                do2ndModel = False
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    p11 = New SLib.InputParameter("Use a secondary model", do2ndModel)
                Else
                    p11 = New SLib.InputParameter("Use the secondary model", do2ndModel)
                End If
            End If

            Dim params() As SLib.InputParameter =
                If(hasSecondary,
                   New SLib.InputParameter() {p0, p1, pSchema, p2, pClampFrom, pClampTo, p3, p4, p6, p9, p10, pMergeEnable, pMergeDateCol, pMergeInstruction, pDebugDiagnostics, p11},
                   New SLib.InputParameter() {p0, p1, pSchema, p2, pClampFrom, pClampTo, p3, p4, p6, p9, p10, pMergeEnable, pMergeDateCol, pMergeInstruction, pDebugDiagnostics})

            ' Optional extra button: "Edit Local Library"
            Dim extraText As String = Nothing
            Dim extraAction As System.Action = Nothing
            Dim closeAfterExtra As Boolean = False

            If Not String.IsNullOrWhiteSpace(localPath) Then
                extraText = "Edit Local Library"
                extraAction =
                        Sub()
                            Try
                                ' Create file with sample content if it doesn't exist or contains only whitespace
                                Dim needsSampleContent As Boolean = False
                                If Not File.Exists(localPath) Then
                                    needsSampleContent = True
                                Else
                                    Try
                                        Dim content As String = File.ReadAllText(localPath, System.Text.Encoding.UTF8)
                                        needsSampleContent = String.IsNullOrWhiteSpace(content)
                                    Catch
                                        needsSampleContent = True
                                    End Try
                                End If

                                If needsSampleContent Then
                                    Try
                                        File.WriteAllText(localPath,
                                            "; Red Ink Fact Extractor - Local Library" & vbCrLf &
                                            "; Format: Title | Instruction | SchemaSpec | MergeEnable | MergeDateCol | MergeInstruction" & vbCrLf &
                                            "; " & vbCrLf &
                                            "; SchemaSpec types: text, number, integer, decimal, date, datetime, other" & vbCrLf &
                                            "; Use * after type to mark preferred sort column (e.g., date*)" & vbCrLf &
                                            "; Lines starting with ; are comments" & vbCrLf &
                                            vbCrLf &
                                            "Contract Summary|Extract the key contract metadata including parties, dates, and financial terms.|Contract Title:text; Party A:text; Party B:text; Effective Date:date*; Expiration Date:date; Contract Value:decimal; Currency:text; Governing Law:text|False|0|" & vbCrLf & vbCrLf &
                                            "Contract Obligations|Extract all contractual obligations, deadlines, and responsible parties.|Obligation:text; Responsible Party:text; Due Date:date*; Frequency:text; Penalty Clause:text|True|3|Merge obligations with the same due date and responsible party" & vbCrLf & vbCrLf &
                                            "Payment Schedule|Extract payment milestones, amounts, and due dates from the contract.|Milestone:text; Payment Date:date*; Amount:decimal; Currency:text; Payment Terms:text; Status:text|True|2|Consolidate payments scheduled for the same date" & vbCrLf & vbCrLf &
                                            "Termination Clauses|Identify all termination and exit provisions.|Clause Type:text; Trigger Condition:text; Notice Period:text; Effective Date:date*; Consequences:text|False|0|" & vbCrLf & vbCrLf &
                                            "Renewal Terms|Extract automatic renewal and extension provisions.|Renewal Type:text; Renewal Date:date*; Duration:text; Notice Deadline:date; Conditions:text|False|0|" & vbCrLf,
                                            System.Text.Encoding.UTF8)
                                    Catch ex As Exception
                                        ShowCustomMessageBox($"Tried to create a sample file but could not: {ex.Message}")
                                        Return
                                    End Try
                                End If

                                ' Open the local library in the editor
                                SLib.ShowTextFileEditor(localPath, $"{AN} Local Library '{localPath}':", False, _context)
                            Catch ex As Exception
                                SLib.ShowCustomMessageBox("Error while opening the local library:" & vbCrLf & ex.Message)
                                Exit Sub
                            End Try

                            ' Inform the user about activation timing
                            SLib.ShowCustomMessageBox("Any changes to the local library will only be active the next time this feature is called up.")
                        End Sub
            End If

            If ShowCustomVariableInputForm("Note: A more sophisticated version of this feature including the possibility to sort and filter columns, to redo failed files and to get clickable file names is available in the Excel add-in as 'Data Extractor'." & vbCrLf & vbCrLf & "Please set the extraction parameters:", AN & " Tabular Overview", params, extraButtonText:=extraText,
                                                                                                                            extraButtonAction:=extraAction,
                                                                                                                            CloseAfterExtra:=closeAfterExtra) = False Then Return

            Dim chosenPreparedDisplay = System.Convert.ToString(params(0).Value)
            manualInstruction = System.Convert.ToString(params(1).Value)

            ' Intercept the builder sentinel: launch wizard, then re-show the main dialog
            ' Only trigger if the user did NOT provide a manual instruction override
            If String.Equals(chosenPreparedDisplay, TabularBuilderSentinel, StringComparison.Ordinal) AndAlso
               String.IsNullOrWhiteSpace(manualInstruction) Then
                Dim builderResult = Await TabularRunLibraryEntryBuilderAsync(localPath)
                If builderResult Then
                    ShowCustomMessageBox("The new library entry has been saved." & vbCrLf &
                                         "Please re-open the Tabular Overview to use it.")
                End If
                Return
            End If

            manualSchemaText = System.Convert.ToString(params(2).Value)
            dateColumnsText = System.Convert.ToString(params(3).Value)
            clampFrom = System.Convert.ToString(params(4).Value)
            clampTo = System.Convert.ToString(params(5).Value)
            sortColumnText = System.Convert.ToString(params(6).Value)
            Dim chosenSortDisplay = System.Convert.ToString(params(7).Value)
            Select Case (If(chosenSortDisplay, "").Trim().ToUpperInvariant())
                Case "DESC", "DESCENDING" : sortDirection = "DESC"
                Case Else : sortDirection = "ASC"
            End Select
            Try : doOcr = System.Convert.ToBoolean(params(8).Value) : Catch : doOcr = False : End Try
            UserOutputLanguage = System.Convert.ToString(params(9).Value)
            dateOutputFormat = System.Convert.ToString(params(10).Value)

            Try : mergeRowsViaLlm = System.Convert.ToBoolean(params(11).Value) : Catch : mergeRowsViaLlm = defaultMergeEnable : End Try
            Dim mergeDateColRaw = System.Convert.ToString(params(12).Value)
            Dim tmpMergeCol As Integer
            If Integer.TryParse(If(mergeDateColRaw, "").Trim(), tmpMergeCol) AndAlso tmpMergeCol > 0 Then
                mergeDateColumn = tmpMergeCol
            Else
                mergeDateColumn = 0
            End If
            mergeInstruction = System.Convert.ToString(params(13).Value)
            Dim debugDiagnostics As System.Boolean = False
            Try : debugDiagnostics = System.Convert.ToBoolean(params(14).Value) : Catch : debugDiagnostics = False : End Try

            If hasSecondary Then
                Try : do2ndModel = System.Convert.ToBoolean(params(15).Value) : Catch : do2ndModel = False : End Try
            End If

            Try : My.Settings.Tabular_ManualInstruction = manualInstruction : Catch : End Try
            Try : My.Settings.Tabular_ManualSchema = manualSchemaText : Catch : End Try
            Try : My.Settings.Tabular_DateColumns = dateColumnsText : Catch : End Try
            Try : My.Settings.Tabular_SortColumn = sortColumnText : Catch : End Try
            Try : My.Settings.Tabular_SortDirection = sortDirection : Catch : End Try
            Try : My.Settings.Tabular_DoOcr = doOcr : Catch : End Try
            Try : My.Settings.Tabular_DateClampFrom = clampFrom : Catch : End Try
            Try : My.Settings.Tabular_DateClampTo = clampTo : Catch : End Try
            Try : My.Settings.Tabular_OutputLanguage = UserOutputLanguage : Catch : End Try
            Try : My.Settings.Tabular_DateOutputFormat = dateOutputFormat : Catch : End Try
            Try : My.Settings.Tabular_MergeEnable = mergeRowsViaLlm : Catch : End Try
            Try : My.Settings.Tabular_MergeDateColumn = If(mergeDateColumn > 0, mergeDateColumn.ToString(), "") : Catch : End Try
            Try : My.Settings.Tabular_MergeInstruction = mergeInstruction : Catch : End Try
            Try : My.Settings.Save() : Catch : End Try

            ' Determine manual override / prepared missing flags
            Dim preparedInstruction As String = Nothing
            If chosenPreparedDisplay IsNot Nothing Then displayToInstruction.TryGetValue(chosenPreparedDisplay, preparedInstruction)
            Dim manualOverrides = Not String.IsNullOrWhiteSpace(manualInstruction)
            Dim preparedMissingInstruction = (Not manualOverrides) AndAlso Not String.IsNullOrWhiteSpace(chosenPreparedDisplay) AndAlso String.IsNullOrWhiteSpace(preparedInstruction)
            If manualOverrides Then
                preparedInstruction = Nothing
                chosenPreparedDisplay = Nothing
            End If

            Dim effectiveInstruction = If(manualOverrides,
                                          manualInstruction,
                                          If(preparedMissingInstruction,
                                             "Extract key factual data points from the document.",
                                             preparedInstruction))
            If String.IsNullOrWhiteSpace(effectiveInstruction) Then
                ShowCustomMessageBox("No extraction instruction provided.")
                Return
            End If

            ' Merge resolution
            Dim userRequestedMerge As Boolean = mergeRowsViaLlm
            Dim userProvidedDateColumn As Boolean = (mergeDateColumn > 0)

            If Not userRequestedMerge Then
                mergeRowsViaLlm = False
                mergeDateColumn = 0
                mergeInstruction = ""
            Else
                If userProvidedDateColumn Then
                    ' Manual override active
                Else
                    If Not String.IsNullOrWhiteSpace(chosenPreparedDisplay) Then
                        Dim libEnable As Boolean = False
                        Dim libDateCol As Integer = 0
                        Dim libInstr As String = ""

                        displayToMergeEnable.TryGetValue(chosenPreparedDisplay, libEnable)
                        displayToMergeDateCol.TryGetValue(chosenPreparedDisplay, libDateCol)
                        displayToMergeInstruction.TryGetValue(chosenPreparedDisplay, libInstr)

                        If libEnable AndAlso libDateCol > 0 Then
                            mergeDateColumn = libDateCol
                            If Not String.IsNullOrWhiteSpace(libInstr) Then
                                mergeInstruction = libInstr
                            End If
                            mergeRowsViaLlm = True
                        Else
                            mergeRowsViaLlm = False
                            mergeDateColumn = 0
                            mergeInstruction = ""
                        End If
                    Else
                        mergeRowsViaLlm = False
                        mergeDateColumn = 0
                        mergeInstruction = ""
                    End If
                End If
            End If

            If hasSecondary AndAlso do2ndModel Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The alternate model could not be loaded - aborting.")
                        Return
                    Else
                        useSecondApi = True
                    End If
                ElseIf INI_SecondAPI Then
                    useSecondApi = True
                End If
            End If

            ' Parse date column list
            Dim dateCols As New System.Collections.Generic.List(Of Integer)
            For Each part In dateColumnsText.Split(New Char() {","c, ";"c}, StringSplitOptions.RemoveEmptyEntries)
                Dim n As Integer
                If Integer.TryParse(part.Trim(), n) AndAlso n > 0 Then dateCols.Add(n)
            Next

            ' Sort column handling with "auto" keyword
            Dim sortColumn As Integer = 0
            Dim wantsAutoSort As Boolean = False
            If Not String.IsNullOrWhiteSpace(sortColumnText) Then
                Dim rawSort = sortColumnText.Trim()
                Dim tmpInt As Integer
                If Integer.TryParse(rawSort, tmpInt) AndAlso tmpInt > 0 Then
                    sortColumn = tmpInt
                ElseIf rawSort.Equals("auto", StringComparison.OrdinalIgnoreCase) Then
                    wantsAutoSort = True
                    sortColumn = 0
                End If
            End If

            Globals.ThisAddIn.OtherPrompt = effectiveInstruction
            Globals.ThisAddIn.OutputLanguage = UserOutputLanguage

            Dim fixedSchema As System.Collections.Generic.List(Of ExtractionSchemaColumn) = Nothing
            Dim autoSortColumn As Integer = 0

            ' Manual schema override (auto-detect only if user typed "auto")
            If Not String.IsNullOrWhiteSpace(manualSchemaText) Then
                fixedSchema = ParseUserSchemaSpec(manualSchemaText)
                If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 AndAlso wantsAutoSort Then
                    autoSortColumn = DetectSortColumnFromSpec(manualSchemaText)
                    If autoSortColumn > 0 Then sortColumn = autoSortColumn
                End If
            End If

            ' Prepared schema (only if no manual schema) with optional auto detection
            If (fixedSchema Is Nothing OrElse fixedSchema.Count = 0) AndAlso
               Not manualOverrides AndAlso Not preparedMissingInstruction AndAlso
               Not String.IsNullOrWhiteSpace(chosenPreparedDisplay) AndAlso displayToSchema.ContainsKey(chosenPreparedDisplay) Then

                Dim spec = displayToSchema(chosenPreparedDisplay)
                If Not String.IsNullOrWhiteSpace(spec) Then
                    fixedSchema = ParseUserSchemaSpec(spec)
                    If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 AndAlso wantsAutoSort Then
                        autoSortColumn = DetectSortColumnFromSpec(spec)
                        If autoSortColumn > 0 Then sortColumn = autoSortColumn
                    End If
                End If
            End If

            ' AI schema generation (only if still no schema)
            If (fixedSchema Is Nothing OrElse fixedSchema.Count = 0) AndAlso (manualOverrides Or preparedMissingInstruction) AndAlso String.IsNullOrWhiteSpace(manualSchemaText) Then
                Dim aiSchema = Await GenerateSchemaFromAiAsync(AddressOf InterpolateAtRuntime,
                                                               AddressOf LLM,
                                                               useSecondApi, _context)
                If aiSchema Is Nothing OrElse aiSchema.Count = 0 Then
                    ShowCustomMessageBox("AI did not return a schema. Aborting.")
                    Return
                End If
                Dim preview = String.Join(vbCrLf, aiSchema.Select(Function(c, i) (i + 1).ToString() & ". " & c.Name & " (" & c.Type & ")"))
                Dim answerSchema = ShowCustomYesNoBox("AI proposed this schema:" & vbCrLf & vbCrLf & preview & vbCrLf & vbCrLf & "Proceed?", "Use schema", "Abort")
                If answerSchema <> 1 Then
                    ShowCustomMessageBox("Operation cancelled.")
                    Return
                End If
                fixedSchema = aiSchema
                If wantsAutoSort AndAlso sortColumn = 0 Then
                    autoSortColumn = DetectSortColumnFromSpec(String.Join(";", aiSchema.Select(Function(sc) sc.Name & ":" & sc.Type)))
                    If autoSortColumn > 0 Then sortColumn = autoSortColumn
                End If
            End If

            ' Ask user: insert at cursor or new document?
            Dim insertInNewDoc As Boolean = False
            Dim destAnswer = ShowCustomYesNoBox("Where should the table be inserted?",
                                                "At cursor position",
                                                "In a new document")
            If destAnswer = 2 Then
                insertInNewDoc = True
            ElseIf destAnswer <> 1 Then
                ShowCustomMessageBox("Operation cancelled.")
                Return
            End If

            ' ── Unified file/folder selection via drag-and-drop form ──
            DragDropFormLabel = ""
            Dim selectedPath As String = Nothing
            Dim selectedIsDirectory As Boolean = False
            Try
                Using form As New DragDropForm(DragDropMode.FileOrDirectory)
                    Dim __safeDialogOwner525 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                    If If(__safeDialogOwner525 IsNot Nothing, form.ShowDialog(__safeDialogOwner525), form.ShowDialog()) <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(form.SelectedFilePath) Then
                        Return
                    End If
                    selectedPath = form.SelectedFilePath.Trim()
                    selectedIsDirectory = form.IsDirectory
                End Using
            Catch ex As Exception
                ShowCustomMessageBox("Selection failed: " & ex.Message)
                Return
            End Try

            If Not selectedIsDirectory Then
                ' ── Single file ──
                Try
                    selectedPath = Path.GetFullPath(selectedPath)
                Catch
                End Try

                If Not File.Exists(selectedPath) Then
                    ShowCustomMessageBox($"The file '{selectedPath}' was not found.")
                    Return
                End If

                Dim list As New System.Collections.Generic.List(Of String) From {selectedPath}

                ShowProgressBarInSeparateThread(AN & " Tabular Overview", "Extracting data...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressMax = list.Count
                GlobalProgressValue = 0
                GlobalProgressLabel = "Starting..."

                Dim cancelFunc As Func(Of Boolean) = Function() ProgressBarModule.CancelOperation

                Dim res As FactExtractionAggregateResult = Nothing
                Try
                    res = Await RunFactExtractionAsync(list,
                                                       effectiveInstruction,
                                                       dateCols,
                                                       sortColumn,
                                                       sortDirection,
                                                       doOcr,
                                                       useSecondApi,
                                                       Path.GetDirectoryName(selectedPath),
                                                       AddressOf InterpolateAtRuntime,
                                                       AddressOf LLM,
                                                       AddressOf GetFileContent,
                                                       _context,
                                                       fixedSchema,
                                                       clampFrom,
                                                       clampTo,
                                                       Sub(cur, total, label)
                                                           GlobalProgressValue = cur
                                                           GlobalProgressMax = total
                                                           GlobalProgressLabel = label
                                                       End Sub,
                                                       mergeDateColumn,
                                                       mergeRowsViaLlm,
                                                       mergeInstruction,
                                                       cancellationRequested:=cancelFunc,
                                                       llmWithFileFunc:=Async Function(sys, usr, mdl, tmp, tmo, use2nd, hide, fileObj)
                                                                            Return Await LLM(sys, usr, mdl, tmp, tmo, use2nd, True, "", fileObj)
                                                                        End Function,
                                                       debugDiagnostics:=debugDiagnostics)
                Catch ex As Exception
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("Single-file extraction failed: " & ex.Message)
                    Return
                Finally
                    ProgressBarModule.CancelOperation = True
                End Try

                If res Is Nothing OrElse res.Rows.Count = 0 Then
                    ShowCustomMessageBox("No data extracted." & If(res IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(res.DebugLogPath), vbCrLf & "Diagnostic log: " & res.DebugLogPath, ""))
                    Return
                End If
                InsertResultIntoWordTable(res, selectedPath, dateOutputFormat, insertInNewDoc)
                ShowCustomMessageBox("Tabular overview completed." & If(res.FailedFiles > 0 AndAlso Not System.String.IsNullOrWhiteSpace(res.DebugLogPath), vbCrLf & "Diagnostic log: " & res.DebugLogPath, ""))
            Else
                ' ── Folder (multiple files) ──
                Dim selectedFolder = selectedPath

                ' Determine whether to include subdirectories
                Dim searchOpt As SearchOption = SearchOption.TopDirectoryOnly
                Try
                    Dim subDirs = Directory.GetDirectories(selectedFolder, "*", SearchOption.TopDirectoryOnly)
                    If subDirs IsNot Nothing AndAlso subDirs.Length > 0 Then
                        Dim subAns = ShowCustomYesNoBox(
                            "The selected folder contains " & subDirs.Length.ToString() &
                            " subfolder(s)." & vbCrLf & vbCrLf &
                            "Do you want to include files from subdirectories?",
                            "Yes, include subdirectories",
                            "No, top-level only")
                        If subAns = 1 Then
                            searchOpt = SearchOption.AllDirectories
                        ElseIf subAns <> 2 Then
                            ShowCustomMessageBox("Operation cancelled.")
                            Return
                        End If
                    End If
                Catch
                End Try

                Dim files() As String = {}
                Dim skippedFiles As New System.Collections.Generic.List(Of String)
                Try
                    Dim baseExtensions = {".pdf", ".docx", ".txt", ".rtf", ".ini", ".csv", ".log",
                                           ".json", ".xml", ".html", ".htm", ".md", ".yaml", ".yml",
                                           ".xlsx", ".pptx", ".msg", ".eml",
                                           ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tiff", ".tif", ".webp", ".svg",
                                           ".mp3", ".wav", ".ogg", ".flac", ".m4a", ".aac", ".wma", ".opus",
                                           ".mp4", ".avi", ".mkv", ".mov", ".wmv", ".webm"}
                    Dim allowedExts = If(INI_AllowLegacyDocFiles,
                                               baseExtensions.Concat({".doc"}).ToArray(),
                                               baseExtensions)

                    allowedExts = allowedExts.
                        Where(Function(ext) IsModelCapableForExtension(_context, ext)).ToArray()

                    Dim allFiles = Directory.GetFiles(selectedFolder, "*.*", searchOpt)
                    Dim supported As New System.Collections.Generic.List(Of String)
                    For Each f In allFiles
                        If allowedExts.Contains(Path.GetExtension(f).ToLowerInvariant()) Then
                            supported.Add(f)
                        Else
                            skippedFiles.Add(f)
                        End If
                    Next
                    files = supported.ToArray()
                Catch ex As Exception
                    ShowCustomMessageBox("Failed to enumerate files: " & ex.Message)
                    Return
                End Try

                ' Warn about large file counts when subdirectories are included
                If searchOpt = SearchOption.AllDirectories AndAlso files.Length > 50 Then
                    Dim warnAns = ShowCustomYesNoBox(
                        "Found " & files.Length.ToString() & " files across all subdirectories." &
                        If(skippedFiles.Count > 0, vbCrLf & skippedFiles.Count.ToString() & " unsupported file(s) will be skipped.", "") &
                        vbCrLf & vbCrLf &
                        "This may take a while. Continue?",
                        "Continue",
                        "Cancel")
                    If warnAns <> 1 Then
                        ShowCustomMessageBox("Operation cancelled.")
                        Return
                    End If
                End If

                If files.Length = 0 Then
                    ShowCustomMessageBox("Folder contains no supported files." &
                        If(skippedFiles.Count > 0, vbCrLf & skippedFiles.Count.ToString() & " unsupported file(s) were found.", ""))
                    Return
                End If

                ShowProgressBarInSeparateThread(AN & " Tabular Overview", "Extracting data...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressMax = files.Length
                GlobalProgressValue = 0
                GlobalProgressLabel = "Starting..."

                Dim cancelFunc As Func(Of Boolean) = Function() ProgressBarModule.CancelOperation

                Dim res = Await RunFactExtractionAsync(New System.Collections.Generic.List(Of String)(files),
                                                       effectiveInstruction,
                                                       dateCols,
                                                       sortColumn,
                                                       sortDirection,
                                                       doOcr,
                                                       useSecondApi,
                                                       selectedFolder,
                                                       AddressOf InterpolateAtRuntime,
                                                       AddressOf LLM,
                                                       AddressOf GetFileContent,
                                                       _context,
                                                       fixedSchema,
                                                       clampFrom,
                                                       clampTo,
                                                       Sub(cur, total, label)
                                                           GlobalProgressValue = cur
                                                           GlobalProgressMax = total
                                                           GlobalProgressLabel = label
                                                       End Sub,
                                                       mergeDateColumn,
                                                       mergeRowsViaLlm,
                                                       mergeInstruction,
                                                       cancellationRequested:=cancelFunc,
                                                       llmWithFileFunc:=Async Function(sys, usr, mdl, tmp, tmo, use2nd, hide, fileObj)
                                                                            Return Await LLM(sys, usr, mdl, tmp, tmo, use2nd, True, "", fileObj)
                                                                        End Function,
                                                       debugDiagnostics:=debugDiagnostics)

                Dim wasCancelled As Boolean = ProgressBarModule.CancelOperation
                ProgressBarModule.CancelOperation = True

                If wasCancelled AndAlso res.Rows.Count = 0 AndAlso res.ProcessedFiles = 0 Then
                    ShowCustomMessageBox("Operation cancelled.")
                    Return
                End If

                If res.Rows.Count = 0 Then
                    Dim msg = "No data extracted."
                    If res.FailedFiles > 0 Then msg &= vbCrLf & "Failed files: " & String.Join(", ", res.FailedFileNames)
                    If Not System.String.IsNullOrWhiteSpace(res.DebugLogPath) Then msg &= vbCrLf & "Diagnostic log: " & res.DebugLogPath
                    ShowCustomMessageBox(msg)
                    Return
                End If
                InsertResultIntoWordTable(res, selectedFolder, dateOutputFormat, insertInNewDoc, skippedFiles)
                Dim summary As New System.Text.StringBuilder()
                summary.AppendLine("Processed files: " & res.ProcessedFiles)
                summary.AppendLine("Failed files: " & res.FailedFiles)
                'If res.FailedFiles > 0 Then
                'summary.AppendLine("Failed file names:")
                'summary.AppendLine(String.Join(", ", res.FailedFileNames))
                'End If
                If skippedFiles.Count > 0 Then
                    summary.AppendLine("Skipped (unsupported): " & skippedFiles.Count.ToString())
                    'summary.AppendLine(String.Join(", ", skippedFiles))
                End If
                If res.FailedFiles > 0 AndAlso Not System.String.IsNullOrWhiteSpace(res.DebugLogPath) Then summary.AppendLine("Diagnostic log: " & res.DebugLogPath)
                ShowCustomMessageBox("Tabular overview completed." & vbCrLf & summary.ToString())
            End If

        Catch ex As Exception
            ShowCustomMessageBox("Tabular overview failed: " & ex.Message)
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Runs the interactive Library Entry Builder wizard for Tabular Overview.
    ''' Prompts the user for a plain-language description, calls the LLM to generate a structured
    ''' library entry, presents it for review/edit, and appends the confirmed entry to the local library file.
    ''' </summary>
    ''' <param name="localPath">Resolved path to the local library file.</param>
    ''' <returns><c>True</c> if a new entry was saved; <c>False</c> if cancelled or failed.</returns>
    Private Async Function TabularRunLibraryEntryBuilderAsync(localPath As String) As System.Threading.Tasks.Task(Of Boolean)
        Try
            ' Step 1: Collect the user's natural-language description
            Dim description = ShowCustomInputBox(
                "Describe what you want to extract from your documents." & vbCrLf &
                "For example: ""I need a timeline of all events with dates, parties involved, " &
                "and a brief description"" or ""Extract invoice metadata: number, date, vendor, " &
                "amounts""." & vbCrLf & vbCrLf &
                "The AI will generate a complete library entry (title, instruction, schema, " &
                "and merge settings) from your description.",
                AN & " Library Entry Builder",
                False,
                "", Context:=_context)

            If String.IsNullOrWhiteSpace(description) OrElse description = "ESC" Then Return False

            ' Step 2: Call LLM to generate the structured library entry
            Dim systemPrompt = SP_ExtractBuilder
            Dim userPrompt = description

            ShowProgressBarInSeparateThread(AN & " Library Builder", "Generating library entry...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = 1
            GlobalProgressValue = 0
            GlobalProgressLabel = "AI is designing your extraction entry..."

            Dim jsonResp As String
            Try
                jsonResp = Await LLM(systemPrompt, userPrompt, "", "", 0, False, True)
            Catch ex As Exception
                ProgressBarModule.CancelOperation = True
                ShowCustomMessageBox("AI generation failed: " & ex.Message)
                Return False
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            jsonResp = SharedLibrary.SharedLibrary.WebAgentInterpreter.SanitizeLlmResult(jsonResp)
            If String.IsNullOrWhiteSpace(jsonResp) Then
                ShowCustomMessageBox("AI did not return a result. Please try again with a more detailed description.")
                Return False
            End If

            ' Step 3: Parse the LLM response
            Dim genTitle As String = ""
            Dim genInstruction As String = ""
            Dim genSchema As String = ""
            Dim genMergeEnabled As Boolean = False
            Dim genMergeColumn As Integer = 0
            Dim genMergeInstruction As String = ""

            Try
                Dim jt = Newtonsoft.Json.Linq.JToken.Parse(jsonResp)
                genTitle = If(CStr(jt("title")), "").Trim()
                genInstruction = If(CStr(jt("instruction")), "").Trim()
                genSchema = If(CStr(jt("schema")), "").Trim()
                Dim me_tok = jt("merge_enabled")
                If me_tok IsNot Nothing Then
                    genMergeEnabled = CBool(me_tok)
                End If
                Dim mc_tok = jt("merge_column")
                If mc_tok IsNot Nothing Then
                    Dim mc = CInt(mc_tok)
                    If mc > 0 Then genMergeColumn = mc
                End If
                genMergeInstruction = If(CStr(jt("merge_instruction")), "").Trim()
            Catch ex As Exception
                ShowCustomMessageBox("Could not parse the AI response. Please try again." & vbCrLf &
                                     "Raw response:" & vbCrLf & If(jsonResp.Length > 500, jsonResp.Substring(0, 500) & "...", jsonResp))
                Return False
            End Try

            If String.IsNullOrWhiteSpace(genTitle) AndAlso String.IsNullOrWhiteSpace(genInstruction) Then
                ShowCustomMessageBox("AI returned an empty entry. Please try again with a more detailed description.")
                Return False
            End If

            ' Step 4: Present for review and editing
            Dim pTitle As New SLib.InputParameter("Title", genTitle)
            Dim pInstr As New SLib.InputParameter("Extraction instruction", genInstruction)
            Dim pSchemaEdit As New SLib.InputParameter("Schema (Name:type; Name:type*; ...)", genSchema)
            Dim pMerge As New SLib.InputParameter("Enable row merging", genMergeEnabled)
            Dim pMergeCol As New SLib.InputParameter("Merge/group column (1-based, 0=none)", If(genMergeColumn > 0, genMergeColumn.ToString(), "0"))
            Dim pMergeInstr As New SLib.InputParameter("Merge instruction (optional)", genMergeInstruction)

            Dim reviewParams() As SLib.InputParameter = {pTitle, pInstr, pSchemaEdit, pMerge, pMergeCol, pMergeInstr}

            If ShowCustomVariableInputForm(
                "Review and edit the AI-generated library entry. Press OK to save to your local library.",
                AN & " Library Entry Builder - Review",
                reviewParams) = False Then Return False

            ' Step 5: Collect final values
            Dim finalTitle = System.Convert.ToString(reviewParams(0).Value).Trim()
            Dim finalInstruction = System.Convert.ToString(reviewParams(1).Value).Trim()
            Dim finalSchema = System.Convert.ToString(reviewParams(2).Value).Trim()
            Dim finalMergeEnabled As Boolean = False
            Try : finalMergeEnabled = System.Convert.ToBoolean(reviewParams(3).Value) : Catch : End Try
            Dim finalMergeColumn As Integer = 0
            Dim mergeColStr = System.Convert.ToString(reviewParams(4).Value).Trim()
            Dim tmpCol As Integer
            If Integer.TryParse(mergeColStr, tmpCol) AndAlso tmpCol > 0 Then finalMergeColumn = tmpCol
            Dim finalMergeInstr = System.Convert.ToString(reviewParams(5).Value).Trim()

            If String.IsNullOrWhiteSpace(finalTitle) Then
                ShowCustomMessageBox("A title is required. Entry not saved.")
                Return False
            End If

            ' Sanitize pipe characters
            finalTitle = finalTitle.Replace("|"c, "-"c)
            finalInstruction = finalInstruction.Replace("|"c, "-"c)
            finalSchema = finalSchema.Replace("|"c, "-"c)
            finalMergeInstr = finalMergeInstr.Replace("|"c, "-"c)

            ' Step 6: Build the library line and append to local library file
            Dim libraryLine As String =
                finalTitle & "|" &
                finalInstruction & "|" &
                finalSchema & "|" &
                If(finalMergeEnabled, "True", "False") & "|" &
                finalMergeColumn.ToString() & "|" &
                finalMergeInstr

            Try
                Dim dir = Path.GetDirectoryName(localPath)
                If Not String.IsNullOrWhiteSpace(dir) AndAlso Not Directory.Exists(dir) Then
                    Directory.CreateDirectory(dir)
                End If

                Dim existingContent As String = ""
                If File.Exists(localPath) Then
                    Try : existingContent = File.ReadAllText(localPath, System.Text.Encoding.UTF8) : Catch : End Try
                End If

                Dim needsHeader = String.IsNullOrWhiteSpace(existingContent)
                Dim sb As New System.Text.StringBuilder()
                If needsHeader Then
                    sb.AppendLine("; Red Ink Fact Extractor - Local Library")
                    sb.AppendLine("; Format: Title | Instruction | SchemaSpec | MergeEnable | MergeDateCol | MergeInstruction")
                    sb.AppendLine("; Lines starting with ; are comments")
                    sb.AppendLine()
                Else
                    sb.Append(existingContent)
                    If Not existingContent.EndsWith(vbCrLf) AndAlso Not existingContent.EndsWith(vbLf) Then
                        sb.AppendLine()
                    End If
                End If

                sb.AppendLine(libraryLine)
                sb.AppendLine()

                File.WriteAllText(localPath, sb.ToString(), System.Text.Encoding.UTF8)
            Catch ex As Exception
                ShowCustomMessageBox("Failed to save the library entry: " & ex.Message)
                Return False
            End Try

            Return True

        Catch ex As Exception
            ShowCustomMessageBox("Library entry builder failed: " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Inserts extracted fact data into a Word table at the current cursor position or in a new document,
    ''' with formatting, optional date normalization, and appended summary rows.
    ''' </summary>
    Private Sub InsertResultIntoWordTable(res As FactExtractionAggregateResult,
                                          basePath As String,
                                          dateFormat As String,
                                          insertInNewDoc As Boolean,
                                          Optional skippedFiles As System.Collections.Generic.List(Of String) = Nothing)
        Try
            Dim cols = res.Schema.Count
            Dim rows = res.Rows.Count
            If cols <= 0 OrElse rows <= 0 Then
                ShowCustomMessageBox("Nothing to insert.")
                Return
            End If

            Dim normalizedDateFormat As String = TabularNormalizeUserDateFormat(dateFormat)

            Dim wordApp As Microsoft.Office.Interop.Word.Application = Globals.ThisAddIn.Application
            Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
            Dim insertRange As Microsoft.Office.Interop.Word.Range = Nothing

            If insertInNewDoc Then
                doc = wordApp.Documents.Add()
                insertRange = doc.Range(0, 0)
            Else
                doc = wordApp.ActiveDocument
                insertRange = wordApp.Selection.Range
            End If

            ' Summary rows: directory, processed, failed + optional failed file names + optional skipped files
            Dim failedCount = If(res.FailedFileNames IsNot Nothing, res.FailedFileNames.Count, 0)
            Dim skippedCount = If(skippedFiles IsNot Nothing, skippedFiles.Count, 0)
            Dim summaryRowCount = 3
            Dim failedDetailRows = If(failedCount > 0, failedCount, 0)
            Dim skippedHeaderRows = If(skippedCount > 0, 1, 0)
            Dim skippedDetailRows = If(skippedCount > 0, skippedCount, 0)
            Dim totalRows = 1 + rows + summaryRowCount + failedDetailRows + skippedHeaderRows + skippedDetailRows

            Dim tbl As Microsoft.Office.Interop.Word.Table = doc.Tables.Add(insertRange, totalRows, cols)
            tbl.Borders.Enable = True

            ' Set table style to a clean grid
            Try
                tbl.Style = "Table Grid"
            Catch
                ' Style may not be available; borders are already enabled
            End Try

            ' Header row
            For c = 0 To cols - 1
                tbl.Cell(1, c + 1).Range.Text = res.Schema(c).Name
            Next
            Try
                tbl.Rows(1).Range.Font.Bold = CInt(True)
                tbl.Rows(1).Shading.BackgroundPatternColor = WdColor.wdColorGray15
            Catch
            End Try

            ' Data rows
            For r = 0 To rows - 1
                For c = 0 To cols - 1
                    Dim v = res.Rows(r).Values(c)
                    Dim outVal As String = If(v Is Nothing, "", v.ToString())
                    Dim typ = If(res.Schema(c).Type, "").ToLowerInvariant()
                    If Not String.IsNullOrWhiteSpace(normalizedDateFormat) AndAlso (typ = "date" OrElse typ = "datetime") Then
                        Dim dt = FactExtractionService.ParseFlexibleDate(outVal)
                        If dt.HasValue Then
                            Try
                                outVal = dt.Value.ToString(normalizedDateFormat, Globalization.CultureInfo.InvariantCulture)
                            Catch
                            End Try
                        End If
                    End If
                    tbl.Cell(2 + r, c + 1).Range.Text = outVal
                Next
            Next

            ' Summary rows (merged across all columns)
            Dim summaryStartRow = 2 + rows
            Dim summaryTexts() As String = {
                "Directory: " & basePath,
                "Files processed: " & res.ProcessedFiles.ToString(),
                "Files failed: " & res.FailedFiles.ToString()
            }

            For i = 0 To summaryTexts.Length - 1
                Dim rowIdx = summaryStartRow + i
                ' Merge all cells in the summary row into one
                If cols > 1 Then
                    Try
                        tbl.Cell(rowIdx, 1).Merge(tbl.Cell(rowIdx, cols))
                    Catch
                    End Try
                End If
                tbl.Cell(rowIdx, 1).Range.Text = summaryTexts(i)
                Try
                    tbl.Cell(rowIdx, 1).Range.Font.Italic = CInt(True)
                    tbl.Cell(rowIdx, 1).Range.Font.Color = WdColor.wdColorGray50
                Catch
                End Try
            Next

            ' Failed file detail rows
            Dim nextRow = summaryStartRow + summaryRowCount
            If failedCount > 0 Then
                For fi = 0 To res.FailedFileNames.Count - 1
                    Dim rowIdx = nextRow + fi
                    Dim fName = res.FailedFileNames(fi)
                    Dim reason As String = ""
                    If res.FailedFileReasons IsNot Nothing Then res.FailedFileReasons.TryGetValue(fName, reason)
                    Dim fPath As String = ""
                    If res.FailedFilePaths IsNot Nothing Then res.FailedFilePaths.TryGetValue(fName, fPath)
                    Dim cellText = If(Not String.IsNullOrWhiteSpace(fPath), fPath, fName) &
                                   If(String.IsNullOrWhiteSpace(reason), "", " [" & reason & "]")

                    ' Merge all cells in the row into one
                    If cols > 1 Then
                        Try
                            tbl.Cell(rowIdx, 1).Merge(tbl.Cell(rowIdx, cols))
                        Catch
                        End Try
                    End If
                    tbl.Cell(rowIdx, 1).Range.Text = cellText
                    Try
                        tbl.Cell(rowIdx, 1).Range.Font.Italic = CInt(True)
                        tbl.Cell(rowIdx, 1).Range.Font.Color = WdColor.wdColorDarkRed
                    Catch
                    End Try
                Next
                nextRow += failedCount
            End If

            ' Skipped (unsupported) file rows
            If skippedCount > 0 Then
                ' Header row for skipped section
                Dim skippedHeaderRow = nextRow
                If cols > 1 Then
                    Try
                        tbl.Cell(skippedHeaderRow, 1).Merge(tbl.Cell(skippedHeaderRow, cols))
                    Catch
                    End Try
                End If
                tbl.Cell(skippedHeaderRow, 1).Range.Text = "Skipped (unsupported): " & skippedCount.ToString()
                Try
                    tbl.Cell(skippedHeaderRow, 1).Range.Font.Italic = CInt(True)
                    tbl.Cell(skippedHeaderRow, 1).Range.Font.Bold = CInt(True)
                    tbl.Cell(skippedHeaderRow, 1).Range.Font.Color = WdColor.wdColorGray50
                Catch
                End Try
                nextRow += 1

                ' Individual skipped file rows
                For si = 0 To skippedFiles.Count - 1
                    Dim rowIdx = nextRow + si
                    If cols > 1 Then
                        Try
                            tbl.Cell(rowIdx, 1).Merge(tbl.Cell(rowIdx, cols))
                        Catch
                        End Try
                    End If
                    tbl.Cell(rowIdx, 1).Range.Text = skippedFiles(si)
                    Try
                        tbl.Cell(rowIdx, 1).Range.Font.Italic = CInt(True)
                        tbl.Cell(rowIdx, 1).Range.Font.Color = WdColor.wdColorGray50
                    Catch
                    End Try
                Next
            End If

            ' AutoFit the table to the window
            Try
                tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
            Catch
            End Try

            ' Move cursor after the table
            Try
                Dim afterTable As Microsoft.Office.Interop.Word.Range = tbl.Range
                afterTable.Collapse(WdCollapseDirection.wdCollapseEnd)
                afterTable.Select()
            Catch
            End Try

        Catch ex As Exception
            ShowCustomMessageBox("Failed inserting table into Word: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Normalizes a user-supplied date format to use 'M' for months in simple month-year patterns
    ''' when lowercase 'm' would otherwise represent minutes.
    ''' </summary>
    Private Function TabularNormalizeUserDateFormat(fmt As String) As String
        If String.IsNullOrWhiteSpace(fmt) Then Return fmt
        Dim hasUpperM = fmt.IndexOf("M"c) <> -1
        Dim hasLowerM = fmt.IndexOf("m"c) <> -1
        Dim hasHourToken = fmt.IndexOf("H"c) <> -1 OrElse fmt.IndexOf("h"c) <> -1
        Dim hasColon = fmt.Contains(":")
        If Not hasUpperM AndAlso hasLowerM AndAlso Not hasHourToken AndAlso Not hasColon Then
            fmt = New String(fmt.Select(Function(ch) If(ch = "m"c, "M"c, ch)).ToArray())
        End If
        Return fmt
    End Function

    ''' <summary>
    ''' Ensures a display string is unique by appending a numeric suffix if a collision exists.
    ''' </summary>
    Private Function TabularMakeUniqueDisplay(baseText As String, existing As ICollection(Of String)) As String
        If existing Is Nothing OrElse existing.Contains(baseText) = False Then
            Return baseText
        End If
        Dim i As Integer = 2
        While existing.Contains(baseText & " [" & i & "]")
            i += 1
        End While
        Return baseText & " [" & i & "]"
    End Function

End Class
