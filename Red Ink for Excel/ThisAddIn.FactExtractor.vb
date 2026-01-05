' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FactExtractor.vb
' Purpose:
'   Orchestrates fact extraction from one or multiple documents into Excel.
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
'   - Execution: Single-file or folder batch; progress via global progress variables.
'   - Result Insertion: Headers, rows, optional date formatting (date/datetime only), wrapping,
'       bounded auto-fit, and appended summary rows.
'   - Cleanup: Restores original model configuration if a secondary model was temporarily loaded.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports SharedLibrary.FactExtractionService

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Main entry point for fact extraction. Collects user parameters, resolves instruction/schema,
    ''' optional secondary model usage, executes extraction (single or multi-file) and inserts results into Excel.
    ''' </summary>
    Public Async Sub FactExtraction()

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
                            Dim unique = MakeUniqueDisplay(display, displayToInstruction.Keys)

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

            Try : defaultManual = System.Convert.ToString(My.Settings.Extractor_ManualInstruction) : Catch : End Try
            Try : defaultManualSchema = System.Convert.ToString(My.Settings.Extractor_ManualSchema) : Catch : End Try
            Try : defaultDateCols = System.Convert.ToString(My.Settings.Extractor_DateColumns) : Catch : End Try
            Try : defaultSortCol = System.Convert.ToString(My.Settings.Extractor_SortColumn) : Catch : End Try
            Try : defaultSortDir = System.Convert.ToString(My.Settings.Extractor_SortDirection) : Catch : End Try
            Try : defaultClampFrom = System.Convert.ToString(My.Settings.Extractor_DateClampFrom) : Catch : End Try
            Try : defaultClampTo = System.Convert.ToString(My.Settings.Extractor_DateClampTo) : Catch : End Try
            Try : defaultOutputLanguage = System.Convert.ToString(My.Settings.Extractor_OutputLanguage) : Catch : End Try
            Try : defaultDoOcr = My.Settings.Extractor_DoOcr : Catch : End Try
            Try : defaultDateOutputFormat = System.Convert.ToString(My.Settings.Extractor_DateOutputFormat) : Catch : End Try
            Try : defaultMergeEnable = My.Settings.Extractor_MergeEnable : Catch : defaultMergeEnable = False : End Try
            Try
                Dim tmp = System.Convert.ToString(My.Settings.Extractor_MergeDateColumn)
                Dim n As Integer
                If Integer.TryParse(tmp, n) AndAlso n > 0 Then defaultMergeDateColumn = n
            Catch
            End Try
            Try : defaultMergeInstruction = System.Convert.ToString(My.Settings.Extractor_MergeInstruction) : Catch : End Try

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
            Dim multipleFiles As Boolean = False
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
            ' Move clamp inputs directly after date column indices
            Dim pClampFrom As New SLib.InputParameter("Only dates on or later", clampFrom)
            Dim pClampTo As New SLib.InputParameter("Only dates up and until", clampTo)
            Dim p3 As New SLib.InputParameter("Column to sort (1-based, auto=*, empty=none)", sortColumnText)
            Dim displaySortDirection = If(sortDirection = "DESC", "Descending", "Ascending")
            Dim p4 As New SLib.InputParameter("Sort direction", displaySortDirection)
            p4.Options = New System.Collections.Generic.List(Of String) From {"Ascending", "Descending"}
            Dim p5 As New SLib.InputParameter("Multiple files", multipleFiles)
            Dim p6 As New SLib.InputParameter("Do OCR if needed (PDFs)", doOcr)
            Dim p9 As New SLib.InputParameter("Output language", UserOutputLanguage)
            Dim p10 As New SLib.InputParameter("Date format (e.g., yyyy-MM-dd; empty=default)", dateOutputFormat)
            Dim pMergeEnable As New SLib.InputParameter("Permit row merging (if requested)", mergeRowsViaLlm)
            Dim pMergeDateCol As New SLib.InputParameter("Column to merge/group on (1-based)", If(mergeDateColumn <= 0, "", mergeDateColumn.ToString()))
            Dim pMergeInstruction As New SLib.InputParameter("Additional merge instructions (optional, overrides)", mergeInstruction)

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
                   New SLib.InputParameter() {p0, p1, pSchema, p2, pClampFrom, pClampTo, p3, p4, p5, p6, p9, p10, pMergeEnable, pMergeDateCol, pMergeInstruction, p11},
                   New SLib.InputParameter() {p0, p1, pSchema, p2, pClampFrom, pClampTo, p3, p4, p5, p6, p9, p10, pMergeEnable, pMergeDateCol, pMergeInstruction})

            ' Optional extra button: “Edit Local Library”
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

            If ShowCustomVariableInputForm("Please set the extraction parameters:", AN & " AI Data Extraction", params, extraButtonText:=extraText,
                                                                                                                            extraButtonAction:=extraAction,
                                                                                                                            CloseAfterExtra:=closeAfterExtra) = False Then Return

            Dim chosenPreparedDisplay = System.Convert.ToString(params(0).Value)
            manualInstruction = System.Convert.ToString(params(1).Value)
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
            Try : multipleFiles = System.Convert.ToBoolean(params(8).Value) : Catch : multipleFiles = False : End Try
            Try : doOcr = System.Convert.ToBoolean(params(9).Value) : Catch : doOcr = False : End Try
            UserOutputLanguage = System.Convert.ToString(params(10).Value)
            dateOutputFormat = System.Convert.ToString(params(11).Value)

            Try : mergeRowsViaLlm = System.Convert.ToBoolean(params(12).Value) : Catch : mergeRowsViaLlm = defaultMergeEnable : End Try
            Dim mergeDateColRaw = System.Convert.ToString(params(13).Value)
            Dim tmpMergeCol As Integer
            If Integer.TryParse(If(mergeDateColRaw, "").Trim(), tmpMergeCol) AndAlso tmpMergeCol > 0 Then
                mergeDateColumn = tmpMergeCol
            Else
                mergeDateColumn = 0
            End If
            mergeInstruction = System.Convert.ToString(params(14).Value)

            If hasSecondary Then
                Try : do2ndModel = System.Convert.ToBoolean(params(15).Value) : Catch : do2ndModel = False : End Try
            End If

            Try : My.Settings.Extractor_ManualInstruction = manualInstruction : Catch : End Try
            Try : My.Settings.Extractor_ManualSchema = manualSchemaText : Catch : End Try
            Try : My.Settings.Extractor_DateColumns = dateColumnsText : Catch : End Try
            Try : My.Settings.Extractor_SortColumn = sortColumnText : Catch : End Try
            Try : My.Settings.Extractor_SortDirection = sortDirection : Catch : End Try
            Try : My.Settings.Extractor_DoOcr = doOcr : Catch : End Try
            Try : My.Settings.Extractor_DateClampFrom = clampFrom : Catch : End Try
            Try : My.Settings.Extractor_DateClampTo = clampTo : Catch : End Try
            Try : My.Settings.Extractor_OutputLanguage = UserOutputLanguage : Catch : End Try
            Try : My.Settings.Extractor_DateOutputFormat = dateOutputFormat : Catch : End Try
            Try : My.Settings.Extractor_MergeEnable = mergeRowsViaLlm : Catch : End Try
            Try : My.Settings.Extractor_MergeDateColumn = If(mergeDateColumn > 0, mergeDateColumn.ToString(), "") : Catch : End Try
            Try : My.Settings.Extractor_MergeInstruction = mergeInstruction : Catch : End Try
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

            ' Rule summary:
            ' Rule summary:
            ' 1. Checkbox (mergeRowsViaLlm) must be True to allow any merging.
            ' 2. If user supplied a merge column (>0) AND checkbox True => manual override (use user's column & instruction).
            ' 3. If user did NOT supply a merge column (>0) but checkbox True:
            '       Use library entry ONLY if library MergeEnabled=True AND library MergeDateColumn>0 (treated as generic key column).
            '       Otherwise disable merging.
            ' 4. If checkbox False => merging disabled.
            ' 5. An instruction alone without a merge key column never triggers merging.

            Dim userRequestedMerge As Boolean = mergeRowsViaLlm
            Dim userProvidedDateColumn As Boolean = (mergeDateColumn > 0)

            If Not userRequestedMerge Then
                ' User did not permit merging
                mergeRowsViaLlm = False
                mergeDateColumn = 0
                mergeInstruction = ""
            Else
                If userProvidedDateColumn Then
                    ' Manual override active: keep user-specified date column & instruction (even if blank)
                Else
                    ' No manual date column; attempt to adopt from library entry
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
                            ' Library insufficient: disable merging
                            mergeRowsViaLlm = False
                            mergeDateColumn = 0
                            mergeInstruction = ""
                        End If
                    Else
                        ' No library entry selected and no manual date column => disable merging
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
                ' If user requested auto and schema just generated, attempt detection
                If wantsAutoSort AndAlso sortColumn = 0 Then
                    autoSortColumn = DetectSortColumnFromSpec(String.Join(";", aiSchema.Select(Function(sc) sc.Name & ":" & sc.Type)))
                    If autoSortColumn > 0 Then sortColumn = autoSortColumn
                End If
            End If

            Dim sheet As Worksheet = CType(Application.ActiveSheet, Worksheet)
            Dim startRow As Integer = 1
            Dim overwriteAtA1 As Boolean = True
            Dim mustPrompt As Boolean = False
            Try
                Dim used = sheet.UsedRange
                If used IsNot Nothing AndAlso used.Rows.Count > 0 AndAlso used.Columns.Count > 0 Then
                    Dim anyValue As Boolean = False
                    Dim vals = used.Value2
                    If vals IsNot Nothing Then
                        If TypeOf vals Is Object(,) Then
                            For Each v In CType(vals, Object(,))
                                If v IsNot Nothing AndAlso v.ToString().Trim().Length > 0 Then
                                    anyValue = True : Exit For
                                End If
                            Next
                        Else
                            If vals.ToString().Trim().Length > 0 Then anyValue = True
                        End If
                    End If
                    mustPrompt = anyValue
                End If
            Catch
            End Try

            If mustPrompt Then
                Dim ans = ShowCustomYesNoBox("Worksheet contains existing content." & vbCrLf & vbCrLf &
                                             "Choose insertion mode for fact table:", "Overwrite at A1", "Append below existing")
                If ans = 1 Then
                    overwriteAtA1 = True
                    startRow = 1
                ElseIf ans = 2 Then
                    overwriteAtA1 = False
                    Dim lastUsedRow As Integer = 1
                    Try
                        Dim ur = sheet.UsedRange
                        lastUsedRow = ur.Row + ur.Rows.Count - 1
                        If lastUsedRow < 1 Then lastUsedRow = 1
                    Catch
                        lastUsedRow = 1
                    End Try
                    startRow = lastUsedRow + 2
                Else
                    ShowCustomMessageBox("Operation cancelled.")
                    Return
                End If
            Else
                overwriteAtA1 = True
                startRow = 1
            End If

            If Not multipleFiles Then
                DragDropFormLabel = ""
                Dim filePath = GetFileName()
                If String.IsNullOrWhiteSpace(filePath) Then Return
                Dim list As New System.Collections.Generic.List(Of String) From {filePath}

                ' Initialize progress bar for single-file flow
                ShowProgressBarInSeparateThread(AN & " Data Extraction", "Extracting data...")
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
                                                       Path.GetDirectoryName(filePath),
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
                                                       cancellationRequested:=cancelFunc)
                Catch ex As Exception
                    ProgressBarModule.CancelOperation = True
                    ShowCustomMessageBox("Single-file extraction failed: " & ex.Message)
                    Return
                Finally
                    ProgressBarModule.CancelOperation = True
                End Try

                If res Is Nothing OrElse res.Rows.Count = 0 Then
                    ShowCustomMessageBox("No data extracted.")
                    Return
                End If
                InsertResultIntoExcel(res, filePath, dateOutputFormat, startRow, overwriteAtA1)
                ShowCustomMessageBox("Data extraction completed.")
            Else
                Dim selectedFolder As String = Nothing
                Try
                    Using dlg As New FolderBrowserDialog()
                        dlg.Description = "Select folder with files to extract data from"
                        dlg.ShowNewFolderButton = False
                        If dlg.ShowDialog() <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                            ShowCustomMessageBox("No folder selected.")
                            Return
                        End If
                        selectedFolder = dlg.SelectedPath
                    End Using
                Catch ex As Exception
                    ShowCustomMessageBox("Folder selection failed: " & ex.Message)
                    Return
                End Try
                Dim files() As String = {}
                Try
                    files = Directory.GetFiles(selectedFolder, "*.*", SearchOption.TopDirectoryOnly).
                        Where(Function(f) {".pdf", ".docx", ".doc", ".txt", ".rtf", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".ht"}.Contains(Path.GetExtension(f).ToLowerInvariant())).ToArray()
                Catch ex As Exception
                    ShowCustomMessageBox("Failed to enumerate files: " & ex.Message)
                    Return
                End Try
                If files.Length = 0 Then
                    ShowCustomMessageBox("Folder contains no supported files.")
                    Return
                End If
                ShowProgressBarInSeparateThread(AN & " Data Extraction", "Extracting data...")
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
                                                       cancellationRequested:=cancelFunc)

                ' Capture cancellation state before closing progress bar
                Dim wasCancelled As Boolean = ProgressBarModule.CancelOperation
                ProgressBarModule.CancelOperation = True

                If wasCancelled AndAlso res.Rows.Count = 0 AndAlso res.ProcessedFiles = 0 Then
                    ShowCustomMessageBox("Operation cancelled.")
                    Return
                End If

                If res.Rows.Count = 0 Then
                    Dim msg = "No data extracted."
                    If res.FailedFiles > 0 Then msg &= vbCrLf & "Failed files: " & String.Join(", ", res.FailedFileNames)
                    ShowCustomMessageBox(msg)
                    Return
                End If
                InsertResultIntoExcel(res, selectedFolder, dateOutputFormat, startRow, overwriteAtA1)
                Dim summary As New System.Text.StringBuilder()
                summary.AppendLine("Processed files: " & res.ProcessedFiles)
                summary.AppendLine("Failed files: " & res.FailedFiles)
                If res.FailedFiles > 0 Then
                    summary.AppendLine("Failed file names:")
                    summary.AppendLine(String.Join(", ", res.FailedFileNames))
                End If
                ShowCustomMessageBox("Data extraction completed." & vbCrLf & summary.ToString())
            End If

        Catch ex As Exception
            ShowCustomMessageBox("Data extraction failed: " & ex.Message)
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                RestoreDefaults(_context, originalConfig)
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Inserts extracted fact data and summary information into the active worksheet with formatting,
    ''' optional date normalization and column width adjustments.
    ''' </summary>
    Private Sub InsertResultIntoExcel(res As FactExtractionAggregateResult,
                                      basePath As String,
                                      dateFormat As String,
                                      startRow As Integer,
                                      overwrite As Boolean)
        Try
            Dim sheet As Worksheet = CType(Application.ActiveSheet, Worksheet)
            Dim cols = res.Schema.Count
            Dim rows = res.Rows.Count
            If cols <= 0 OrElse rows <= 0 Then
                ShowCustomMessageBox("Nothing to insert.")
                Return
            End If

            Dim normalizedDateFormat As String = NormalizeUserDateFormat(dateFormat)

            Dim startCol As Integer = 1
            Dim summaryRows = If(res.FailedFiles > 0, 4, 3)
            Dim endRowEstimate = startRow + rows + 1 + summaryRows

            If overwrite Then
                Try
                    Dim clearRange As Range = sheet.Range(sheet.Cells(startRow, startCol),
                                                          sheet.Cells(endRowEstimate, startCol + cols - 1))
                    clearRange.Clear()
                Catch
                End Try
            End If

            For c = 0 To cols - 1
                sheet.Cells(startRow, startCol + c).Value2 = res.Schema(c).Name
            Next
            Dim headerRange As Range = sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(startRow, startCol + cols - 1))
            headerRange.Font.Bold = True
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignLeft

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
                    sheet.Cells(startRow + 1 + r, startCol + c).Value2 = outVal
                Next
            Next
            Dim dataRange As Range = sheet.Range(sheet.Cells(startRow + 1, startCol), sheet.Cells(startRow + rows, startCol + cols - 1))
            dataRange.HorizontalAlignment = XlHAlign.xlHAlignLeft

            Dim summaryRow = startRow + rows + 2
            sheet.Cells(summaryRow, startCol).Value2 = "Directory:"
            sheet.Cells(summaryRow, startCol + 1).Value2 = basePath
            sheet.Cells(summaryRow + 1, startCol).Value2 = "Files processed:"
            sheet.Cells(summaryRow + 1, startCol + 1).Value2 = res.ProcessedFiles
            sheet.Cells(summaryRow + 2, startCol).Value2 = "Files failed:"
            sheet.Cells(summaryRow + 2, startCol + 1).Value2 = res.FailedFiles
            If res.FailedFiles > 0 Then
                sheet.Cells(summaryRow + 3, startCol).Value2 = "Failed file list:"
                sheet.Cells(summaryRow + 3, startCol + 1).Value2 = String.Join(", ", res.FailedFileNames)
            End If
            Dim summaryLastRow = summaryRow + If(res.FailedFiles > 0, 3, 2)
            Dim summaryRange As Range = sheet.Range(sheet.Cells(summaryRow, startCol), sheet.Cells(summaryLastRow, startCol + 1))
            summaryRange.Font.Italic = True
            summaryRange.HorizontalAlignment = XlHAlign.xlHAlignLeft

            headerRange.VerticalAlignment = XlVAlign.xlVAlignTop
            dataRange.VerticalAlignment = XlVAlign.xlVAlignTop
            summaryRange.VerticalAlignment = XlVAlign.xlVAlignTop

            Dim lastDataRow = startRow + rows
            For c = 0 To cols - 1
                Dim typ = If(res.Schema(c).Type, "").ToLowerInvariant()
                If typ = "text" OrElse typ = "other" OrElse typ = "date" OrElse typ = "datetime" Then
                    Dim colRange As Range = sheet.Range(sheet.Cells(startRow, startCol + c), sheet.Cells(lastDataRow, startCol + c))
                    colRange.WrapText = True
                End If
            Next
            summaryRange.WrapText = True

            Dim fullUsedRange As Range = sheet.Range(sheet.Cells(startRow, startCol), sheet.Cells(summaryLastRow, startCol + cols - 1))
            fullUsedRange.EntireColumn.AutoFit()

            For c = 0 To cols - 1
                Dim typ = If(res.Schema(c).Type, "").ToLowerInvariant()
                If typ = "text" OrElse typ = "other" Then
                    Dim columnRange As Range = sheet.Columns(startCol + c)
                    Try
                        If columnRange.ColumnWidth < 30 Then columnRange.ColumnWidth = 30
                        If columnRange.ColumnWidth > 120 Then columnRange.ColumnWidth = 120
                    Catch
                    End Try
                End If
            Next

            headerRange.Rows.AutoFit()
            dataRange.Rows.AutoFit()
            summaryRange.Rows.AutoFit()

        Catch ex As Exception
            ShowCustomMessageBox("Failed inserting into Excel: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Normalizes a user-supplied date format to use 'M' for months in simple month-year patterns
    ''' when lowercase 'm' would otherwise represent minutes.
    ''' </summary>
    Private Function NormalizeUserDateFormat(fmt As String) As String
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
    Private Function MakeUniqueDisplay(baseText As String, existing As ICollection(Of String)) As String
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

' =================================================================================================
' FACT EXTRACTOR CONFIGURATION AND SYNTAX REFERENCE
' -------------------------------------------------------------------------------------------------
' This add-in supports a "prepared instruction set" loaded from one or more text files and optional
' manual overrides entered in the parameter dialog. It extracts facts into Excel following a
' schema-first workflow with explicit date handling, clamping, sorting, and formatting.
'
' -------------------------------------------------------------------------------------------------
' PREPARED INSTRUCTION FILES
' -------------------------------------------------------------------------------------------------
' Location:
'   - Local path:    resolved from INI_ExtractorPathLocal (environment variables expanded)
'   - Global path:   resolved from INI_ExtractorPath (environment variables expanded)
'   - Each location may be a folder (all *.txt files are read, top-level only) or a single file.
'
' File format (one entry per line, pipe-separated):
'   Title | Instruction | SchemaSpec
'
'   - Title (required):
'       Display name shown in the dropdown. If the file is from the local path, " (local)" is appended.
'       If duplicates occur, "[n]" is added to make it unique.
'
'   - Instruction (optional):
'       Natural-language instruction describing what to extract (e.g., "Extract invoice number,
'       vendor, dates, and totals").
'
'   - SchemaSpec (optional):
'       Schema definition for the resulting table (see "Schema specification" below).
'
' Notes:
'   - Blank lines are ignored.
'   - Lines starting with ";" are comments and ignored.
'   - If a prepared entry has a title but no instruction, and no manual instruction is provided,
'     a generic fallback instruction is used ("Extract key factual data points from the document.").
'
' -------------------------------------------------------------------------------------------------
' MANUAL OVERRIDES VS. PREPARED SETS
' -------------------------------------------------------------------------------------------------
' - If "Manual instruction (overrides)" is non-empty, it overrides any prepared set.
' - "Manual schema" is used when provided; otherwise the schema from the prepared entry is used.
' - If neither manual nor prepared schema is available, the add-in can request an AI-generated schema
'   based on the instruction (requires a model and user confirmation).
'
' -------------------------------------------------------------------------------------------------
' SCHEMA SPECIFICATION (semicolon-separated Name[:type][*])
' -------------------------------------------------------------------------------------------------
' Syntax:
'   Name[:type][*]; Name[:type][*]; ...
'
' Components:
'   - Name: required, the column header as it will appear in Excel.
'   - :type: optional, one of the recognized types (case-insensitive):
'       text | number | integer | decimal | date | datetime | other
'     If omitted, type defaults to "text" unless auto-inferred elsewhere.
'   - * (asterisk): optional flag indicating "preferred sort column" when the user selects
'     Sort column = "auto". Only one column should carry "*" for predictable behavior.
'
' Examples:
'   "Event:text; Event Date:date*; Country:text; Report Date:date"
'   "Invoice #:text; Date:date*; Vendor:text; Amount:decimal"
'
' Behavior:
'   - When Sort column = "auto", the first column in the spec marked with "*" is used for sorting.
'   - Columns typed as "date" or "datetime" are treated as date-like and eligible for normalization
'     and formatting. Date normalization/formatting runs only for these types.
'   - The final fact table also appends a "File" column containing the source file name. This column
'     is not part of the schema spec and is added at the end automatically.
'
' -------------------------------------------------------------------------------------------------
' DATE COLUMN INDICES (CSV or semicolon, 1-based)
' -------------------------------------------------------------------------------------------------
' Purpose:
'   Identify which schema columns contain date-like values for normalization, clamping, and improved sorting.
'
' Format:
'   - Comma or semicolon separated integers: "2,5" | "1;3;4" | "7" | "2, 4; 6"
'   - 1-based indexing: 1 = first column of the output table (header row not counted).
'   - Indices refer only to columns defined by the schema. The automatically appended "File" column
'     should not be included.
'
' Validation:
'   - Non-numeric tokens, zeros, negatives, and out-of-range values are ignored.
'   - Duplicates are harmless.
'
' Effects:
'   - Values in the specified columns are normalized as dates during merge.
'   - Date clamp filters (see next section) evaluate these columns; a row is excluded if any listed
'     date column falls outside the specified bounds.
'
' -------------------------------------------------------------------------------------------------
' DATE CLAMP INPUTS ("Only dates on or later" / "Only dates up and until")
' -------------------------------------------------------------------------------------------------
' Supported explicit inputs (case-insensitive month names where applicable):
'   1) yyyy-MM-dd                ISO full date (e.g., 2024-11-24)
'   2) yyyy-MM or yyyy-M         treated as the first day of that month (e.g., 2024-7 => 2024-07-01)
'   3) yyyy                      treated as January 1 of that year (e.g., 2024 => 2024-01-01)
'   4) d.M.yyyy / dd.MM.yyyy     dot day-month-year (e.g., 5.7.2024 / 05.07.2024)
'   5) d.M.yy / dd.MM.yy         dot day-month-two-digit year (e.g., 5.7.24 / 05.07.24)
'                                yy < 50 => 2000–2049; yy >= 50 => 1950–1999
'   6) Full month name + year    InvariantCulture full month name (e.g., "January 2024", "December 2023")
'
' Fallback parsing:
'   If none of the above match, ParseFlexibleDate tries DateTime.TryParse with InvariantCulture,
'   then CurrentCulture. This may allow culture-specific inputs (e.g., 7/5/2024) but is less predictable.
'   Prefer the explicit formats listed above for reliability.
'
' Precision handling:
'   - yyyy        -> coerced to yyyy-01-01
'   - yyyy-MM     -> coerced to yyyy-MM-01
'   - yyyy-MM-dd  -> exact day as provided
'
' Clamp comparison:
'   - Boundaries are inclusive.
'   - Empty "from" or "to" means that boundary is not applied.
'   - Invalid inputs (cannot be parsed) are ignored (no clamp applied on that side).
'   - A row is kept only if every specified date column is within both boundaries.
'
' Recommendations:
'   Use ISO full dates (yyyy-MM-dd) for unambiguous boundaries.
'
' -------------------------------------------------------------------------------------------------
' SORTING
' -------------------------------------------------------------------------------------------------
' Parameter: "Column to sort (1-based, auto=*, empty=none)"
'
' Options:
'   - Empty: no sort applied.
'   - 1-based integer: explicit column in the output table to sort by.
'   - "auto": uses the first column marked with "*" in the schema spec.
'
' Direction:
'   - "Ascending" or "Descending" (display values). Internally mapped to ASC/DESC.
'
' -------------------------------------------------------------------------------------------------
' DATE/TIME OUTPUT FORMAT (Normalization and Formatting Rules)
' -------------------------------------------------------------------------------------------------
' Where formatting applies:
'   - Only columns typed as "date" or "datetime" in the schema are formatted.
'   - If the "Date format" parameter is empty, the normalized date string is kept as-is.
'   - If a non-empty "Date format" is provided, each parsed date is formatted using
'     the specified .NET custom DateTime format with CultureInfo.InvariantCulture.
'
' Accepted .NET custom format tokens (subset commonly used):
'   Day:
'     d      day (1–31)
'     dd     day zero-padded (01–31)
'     ddd    abbreviated weekday (Mon)
'     dddd   full weekday (Monday)
'
'   Month:
'     M      month (1–12)
'     MM     month zero-padded (01–12)
'     MMM    abbreviated month (Jan)
'     MMMM   full month (January)
'     NOTE: Lowercase m means MINUTES, not months.
'
'   Year:
'     yy     two-digit year (24)
'     yyyy   four-digit year (2024)
'
'   Time:
'     H / HH   24-hour clock hour (0–23 / zero-padded)
'     h / hh   12-hour clock hour (1–12 / zero-padded)
'     m / mm   minutes (0–59 / zero-padded)  (Do not use for months!)
'     s / ss   seconds (0–59 / zero-padded)
'     f..fffffff   fractional seconds
'     tt       AM / PM designator
'
'   Time zone / offset (when present in data):
'     z / zz / zzz  offset hours / hours+minutes
'     K             kind / offset / Z
'
' Auto-correction rule (to avoid month/minute confusion):
'   NormalizeUserDateFormat(fmt):
'     If all of the following are true:
'       - The format contains lowercase 'm'
'       - The format does NOT contain uppercase 'M'
'       - The format does NOT contain hour tokens ('H' or 'h')
'       - The format does NOT contain a colon ':'
'     Then every 'm' is converted to 'M' before formatting.
'     Examples auto-corrected:
'       "mmm yy"  -> "MMM yy"
'       "mm yyyy" -> "MM yyyy"
'     Examples NOT auto-corrected (you intended minutes):
'       "HH:mm"   (contains hour token and colon)
'       "h:mm tt" (contains hour token and colon)
'
' Formatting behavior:
'   - Each output value is first parsed with ParseFlexibleDate; if a date is recognized, it is
'     formatted using the provided pattern and InvariantCulture.
'   - If parsing fails or the format string is invalid, formatting is skipped and the original
'     normalized value is left unchanged.
'
' Safe format examples:
'   yyyy-MM-dd
'   dd.MM.yyyy
'   MMM yy
'   MMMM yyyy
'   yyyyMMdd
'   dd MMM yyyy
'   yyyy-MM-dd HH:mm
'
' -------------------------------------------------------------------------------------------------
' OTHER PARAMETERS
' -------------------------------------------------------------------------------------------------
' - Output language:
'     Hint to the model/normalization for language-specific outputs where applicable.
'
' - Multiple files:
'     When enabled, prompts for a folder and processes all supported files in the top directory
'     (.pdf, .docx, .doc, .txt, .rtf, .ini, .csv, .log, .json, .xml, .html, .ht).
'     A progress UI is shown. Summary rows are appended after data.
'
' - Do OCR if needed (PDFs):
'     Attempts OCR for unsearchable PDFs.
'
' - Secondary model:
'     If configured (alternate model path or second API), toggling this will use the secondary
'     model for extraction. If an alternate model requires selection and fails to load, the operation aborts.
'
' -------------------------------------------------------------------------------------------------
' EXCEL INSERTION BEHAVIOR
' -------------------------------------------------------------------------------------------------
' - If the active sheet has content, the user is prompted to overwrite at A1 or append below existing data.
' - Headers use the schema column names; data rows follow.
' - A summary block is appended listing directory, files processed, and failures (optionally names).
' - Text/date/other columns are wrapped and columns auto-fit with width safety caps for text columns.
'
' -------------------------------------------------------------------------------------------------
' ERROR HANDLING
' -------------------------------------------------------------------------------------------------
' - Most parsing/formatting operations are fail-safe and fall back to defaults.
' - If AI schema generation returns nothing or is declined, the operation is cancelled.
' - Exceptions during insertion or file enumeration show a message and abort gracefully.
' =================================================================================================