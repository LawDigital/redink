' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: FactExtractionService.vb
' Purpose: Aggregates structured "fact extraction" results across multiple input files by calling an LLM
'          and parsing a JSON response into a common schema + row set.
'
' Architecture:
'  - Schema Model: `ExtractionSchemaColumn` defines column name and a simple type hint (e.g., text/date/number).
'  - Row Model: `ExtractionRow` stores values aligned to the aggregate schema order.
'  - Parsing: `ParseSingleFileJson` parses LLM JSON payloads with `schema`, `rows`, `file_name`, `notes`.
'    `rows` may be array-of-arrays or array-of-objects (mapped by column names).
'  - Date Handling: `ParseFlexibleDate` and `NormalizeDate` normalize user/LLM-provided date strings.
'    Date columns specified by 1-based indices are normalized on merge and can be range-clamped.
'  - Aggregation: `MergeIntoAggregate` merges per-file schema + rows into a master result and adds a `File` column.
'  - Sorting: `SortAggregate` sorts rows by a 1-based column index using type hints for date-like columns.
'  - Orchestration: `RunFactExtractionAsync` iterates files, loads content, invokes the LLM, merges results,
'    optionally merges grouped rows via an LLM, applies clamps, sorts, and reports progress/cancellation.
'
' External Dependencies:
'  - `Newtonsoft.Json.Linq` for JSON parsing.
'  - `WebAgentInterpreter.SanitizeLlmResult` for response sanitization.
'  - `ISharedContext` for prompt templates (`SP_Extract`, `SP_ExtractSchema`, `SP_MergeDateRows`).
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Text
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary

    ''' <summary>
    ''' Fact extraction orchestration helpers: schema parsing, date normalization, aggregation, sorting, and optional row merging.
    ''' </summary>
    Public Module FactExtractionService

        ''' <summary>
        ''' Defines a single schema column produced/consumed by extraction.
        ''' </summary>
        <DebuggerDisplay("{Name} ({Type})")>
        Public Class ExtractionSchemaColumn
            ''' <summary>
            ''' Column name as used in schema and row mappings.
            ''' </summary>
            Public Property Name As String

            ''' <summary>
            ''' Type hint for comparison/normalization: text | date | datetime | time | number | other.
            ''' </summary>
            Public Property Type As String ' text | date | datetime | time | number | other
        End Class

        ''' <summary>
        ''' Represents one extracted row aligned to the aggregate schema order.
        ''' </summary>
        Public Class ExtractionRow
            ''' <summary>
            ''' Row cell values in schema order.
            ''' </summary>
            Public Property Values As System.Collections.Generic.List(Of Object)
        End Class

        ''' <summary>
        ''' Aggregate extraction result across multiple input files.
        ''' </summary>
        Public Class FactExtractionAggregateResult
            ''' <summary>
            ''' Aggregate schema across all processed files.
            ''' </summary>
            Public Property Schema As System.Collections.Generic.List(Of ExtractionSchemaColumn)

            ''' <summary>
            ''' Aggregate rows across all processed files (aligned to <see cref="Schema"/>).
            ''' </summary>
            Public Property Rows As System.Collections.Generic.List(Of ExtractionRow)

            ''' <summary>
            ''' Collected non-fatal errors encountered during processing.
            ''' </summary>
            Public Property Errors As System.Collections.Generic.List(Of String)

            ''' <summary>
            ''' Count of files successfully processed and merged.
            ''' </summary>
            Public Property ProcessedFiles As Integer

            ''' <summary>
            ''' Count of files that failed (derived from <see cref="FailedFileNames"/>).
            ''' </summary>
            Public Property FailedFiles As Integer

            ''' <summary>
            ''' Basenames of files that failed (missing, empty text, parse failures, LLM failures).
            ''' </summary>
            Public Property FailedFileNames As System.Collections.Generic.List(Of String)

            ''' <summary>
            ''' Source directory associated with the run.
            ''' </summary>
            Public Property SourceDirectory As String
        End Class

        ''' <summary>
        ''' Setting key: manual instruction for extraction run.
        ''' </summary>
        Public Const Setting_ManualInstruction As String = "Extraction_ManualInstruction"

        ''' <summary>
        ''' Setting key: user-specified date columns (1-based index list).
        ''' </summary>
        Public Const Setting_DateColumns As String = "Extraction_DateColumns"

        ''' <summary>
        ''' Setting key: sort column (1-based index).
        ''' </summary>
        Public Const Setting_SortColumn As String = "Extraction_SortColumn"

        ''' <summary>
        ''' Setting key: sort direction (ASC/DESC).
        ''' </summary>
        Public Const Setting_SortDirection As String = "Extraction_SortDirection"

        ''' <summary>
        ''' Setting key: whether OCR should be used when loading file content.
        ''' </summary>
        Public Const Setting_DoOcr As String = "Extraction_DoOcr"

        ''' <summary>
        ''' Setting key: inclusive lower date clamp bound for filtering rows.
        ''' </summary>
        Public Const Setting_DateClampFrom As String = "Extraction_DateClampFrom"

        ''' <summary>
        ''' Setting key: inclusive upper date clamp bound for filtering rows.
        ''' </summary>
        Public Const Setting_DateClampTo As String = "Extraction_DateClampTo"

        ''' <summary>
        ''' Setting key: output language for extraction.
        ''' </summary>
        Public Const Setting_OutputLanguage As String = "Extraction_OutputLanguage"

        ''' <summary>
        ''' Setting key: output format for dates (currently not applied in this module).
        ''' </summary>
        Public Const Setting_DateOutputFormat As String = "Extraction_DateOutputFormat"


        ''' <summary>
        ''' Parses a date string using supported formats, returning <see cref="Date"/> when recognized; otherwise <c>Nothing</c>.
        ''' </summary>
        Public Function ParseFlexibleDate(raw As String) As Date?
            If String.IsNullOrWhiteSpace(raw) Then Return Nothing
            Dim t = raw.Trim()

            Dim mShort = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-([1-9]|1[0-2])$")
            If mShort.Success Then
                Dim yearPart = mShort.Value.Substring(0, 4)
                Dim monthPart = mShort.Groups(2).Value.PadLeft(2, "0"c)
                Return New Date(CInt(yearPart), CInt(monthPart), 1)
            End If

            Dim isoMonth = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])$")
            If isoMonth.Success Then
                Dim yearPart = t.Substring(0, 4)
                Dim monthPart = t.Substring(5, 2)
                Return New Date(CInt(yearPart), CInt(monthPart), 1)
            End If

            Dim isoFull = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$")
            If isoFull.Success Then
                Try : Return Date.ParseExact(t, "yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture) : Catch : End Try
            End If

            Dim dmY = System.Text.RegularExpressions.Regex.Match(t, "^(?<d>[0-3]?\d)\.(?<m>[0-1]?\d)\.(?<y>(19|20)\d{2})$")
            If dmY.Success Then
                Try
                    Return New Date(CInt(dmY.Groups("y").Value), CInt(dmY.Groups("m").Value), CInt(dmY.Groups("d").Value))
                Catch
                End Try
            End If

            Dim dmYY = System.Text.RegularExpressions.Regex.Match(t, "^(?<d>[0-3]?\d)\.(?<m>[0-1]?\d)\.(?<y>\d{2})$")
            If dmYY.Success Then
                Try
                    Dim yy = CInt(dmYY.Groups("y").Value)
                    Dim y = If(yy < 50, 2000 + yy, 1900 + yy)
                    Return New Date(y, CInt(dmYY.Groups("m").Value), CInt(dmYY.Groups("d").Value))
                Catch
                End Try
            End If

            Dim monthYear = System.Text.RegularExpressions.Regex.Match(t, "^(?<m>\p{L}+)\s+(?<y>(19|20)\d{2})$")
            If monthYear.Success Then
                Try
                    Dim y = CInt(monthYear.Groups("y").Value)
                    Dim monthName = monthYear.Groups("m").Value
                    Dim m = DateTime.ParseExact(monthName, "MMMM", Globalization.CultureInfo.InvariantCulture).Month
                    Return New Date(y, m, 1)
                Catch
                End Try
            End If

            Dim yearOnly = System.Text.RegularExpressions.Regex.Match(t, "^(19|20)\d{2}$")
            If yearOnly.Success Then Return New Date(CInt(yearOnly.Value), 1, 1)

            Dim dt As DateTime
            If DateTime.TryParse(t, Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.AllowWhiteSpaces, dt) Then Return dt
            If DateTime.TryParse(t, Globalization.CultureInfo.CurrentCulture, Globalization.DateTimeStyles.AllowWhiteSpaces, dt) Then Return dt
            Return Nothing
        End Function

        ''' <summary>
        ''' Normalizes recognized dates to ISO-like forms depending on the apparent input precision (year, year-month, or full date).
        ''' Unrecognized values are returned unchanged.
        ''' </summary>
        Public Function NormalizeDate(raw As String) As String
            Dim p = ParseFlexibleDate(raw)
            If Not p.HasValue Then Return raw
            Dim t = raw.Trim()
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}$") Then Return p.Value.ToString("yyyy", Globalization.CultureInfo.InvariantCulture)
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}-(0?[1-9]|1[0-2])$") _
               OrElse System.Text.RegularExpressions.Regex.IsMatch(t, "^\p{L}+\s+(19|20)\d{2}$") Then
                Return p.Value.ToString("yyyy-MM", Globalization.CultureInfo.InvariantCulture)
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^[0-3]?\d\.[0-1]?\d\.(\d{2}|\d{4})$") Then Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            If System.Text.RegularExpressions.Regex.IsMatch(t, "^(19|20)\d{2}-(0[1-9]|1[0-2])-(0[1-9]|[12]\d|3[01])$") Then Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
            Return p.Value.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
        End Function

        ''' <summary>
        ''' Converts normalized date strings (yyyy / yyyy-MM / yyyy-MM-dd or other supported formats) into comparable <see cref="Date"/> values.
        ''' </summary>
        Private Function ToComparableDate(normalized As String) As Date?
            If String.IsNullOrWhiteSpace(normalized) Then Return Nothing
            Dim s = normalized.Trim()
            Dim yearOnly = System.Text.RegularExpressions.Regex.Match(s, "^(?<year>(19|20)\d{2})$")
            If yearOnly.Success Then Return New Date(CInt(yearOnly.Groups("year").Value), 1, 1)
            Dim yearMonth = System.Text.RegularExpressions.Regex.Match(s, "^(?<year>(19|20)\d{2})-(?<month>(0[1-9]|1[0-2]))$")
            If yearMonth.Success Then Return New Date(CInt(yearMonth.Groups("year").Value), CInt(yearMonth.Groups("month").Value), 1)
            Return ParseFlexibleDate(s)
        End Function

        ''' <summary>
        ''' Parses a single-file extraction JSON payload into schema, rows, file name and notes.
        ''' Expected top-level keys: <c>schema</c>, <c>rows</c>, <c>file_name</c>, <c>notes</c>.
        ''' </summary>
        Public Function ParseSingleFileJson(json As String) As (schema As System.Collections.Generic.List(Of ExtractionSchemaColumn), rows As System.Collections.Generic.List(Of ExtractionRow), fileName As String, notes As String)
            Dim schema As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            Dim rows As New System.Collections.Generic.List(Of ExtractionRow)
            Dim fileName As String = ""
            Dim notes As String = ""
            If String.IsNullOrWhiteSpace(json) Then Return (schema, rows, fileName, notes)
            Try
                Dim jt = JToken.Parse(json)
                Dim schemaTok = jt("schema")
                If schemaTok IsNot Nothing AndAlso schemaTok.Type = JTokenType.Array Then
                    For Each c In schemaTok
                        Dim name = CStr(c("name"))
                        Dim typ = CStr(c("type"))
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        schema.Add(New ExtractionSchemaColumn With {.Name = name.Trim(), .Type = If(String.IsNullOrWhiteSpace(typ), "text", typ.Trim().ToLowerInvariant())})
                    Next
                End If
                Dim rowsTok = jt("rows")
                If rowsTok IsNot Nothing AndAlso rowsTok.Type = JTokenType.Array Then
                    For Each r In rowsTok
                        Dim er As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                        If r.Type = JTokenType.Array Then
                            For Each v In DirectCast(r, JArray)
                                er.Values.Add(ConvertToken(v))
                            Next
                        ElseIf r.Type = JTokenType.Object Then
                            For Each col In schema
                                er.Values.Add(ConvertToken(r(col.Name)))
                            Next
                        End If
                        If er.Values.Count > 0 Then rows.Add(er)
                    Next
                End If
                fileName = CStr(jt("file_name"))
                notes = CStr(jt("notes"))
            Catch
            End Try
            Return (schema, rows, fileName, notes)
        End Function

        ''' <summary>
        ''' Converts a JSON token into a VB value suitable for storing in <see cref="ExtractionRow.Values"/>.
        ''' </summary>
        Private Function ConvertToken(tok As JToken) As Object
            If tok Is Nothing Then Return ""
            Select Case tok.Type
                Case JTokenType.String : Return CStr(tok)
                Case JTokenType.Integer : Return CInt(tok)
                Case JTokenType.Float : Return CDbl(tok)
                Case JTokenType.Boolean : Return CBool(tok)
                Case Else : Return tok.ToString()
            End Select
        End Function

        ''' <summary>
        ''' Parses a user schema specification string (<c>name[:type][*]</c> tokens separated by <c>;</c>) into a schema list.
        ''' The optional trailing <c>*</c> marks the column as a sort column for <see cref="DetectSortColumnFromSpec"/>.
        ''' </summary>
        Public Function ParseUserSchemaSpec(spec As String) As System.Collections.Generic.List(Of ExtractionSchemaColumn)
            Dim result As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            If String.IsNullOrWhiteSpace(spec) Then Return result
            For Each raw In spec.Split(";"c)
                Dim token = raw.Trim()
                If token.Length = 0 Then Continue For
                Dim namePart = token
                Dim typePart = "text"
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 Then
                    namePart = token.Substring(0, colonIdx).Trim()
                    Dim after = token.Substring(colonIdx + 1).Trim()
                    If after.EndsWith("*") Then after = after.Substring(0, after.Length - 1).Trim()
                    If after.Length > 0 Then typePart = after.ToLowerInvariant()
                ElseIf token.EndsWith("*") Then
                    namePart = token.Substring(0, token.Length - 1).Trim()
                End If
                If String.IsNullOrWhiteSpace(namePart) Then Continue For
                result.Add(New ExtractionSchemaColumn With {.Name = namePart, .Type = typePart})
            Next
            Return result
        End Function

        ''' <summary>
        ''' Detects the 1-based sort column index from a schema specification string by finding a token marked with trailing <c>*</c>.
        ''' Returns 0 when no sort column is marked.
        ''' </summary>
        Public Function DetectSortColumnFromSpec(spec As String) As Integer
            If String.IsNullOrWhiteSpace(spec) Then Return 0
            Dim idx = 0
            For Each raw In spec.Split(";"c)
                Dim token = raw.Trim()
                If token.Length = 0 Then Continue For
                idx += 1
                If token.EndsWith("*") Then Return idx
                Dim colonIdx = token.IndexOf(":"c)
                If colonIdx >= 0 Then
                    Dim after = token.Substring(colonIdx + 1).Trim()
                    If after.EndsWith("*") Then Return idx
                End If
            Next
            Return 0
        End Function

        ''' <summary>
        ''' Calls the LLM to generate a schema-only JSON payload and returns the parsed schema list.
        ''' </summary>
        Public Async Function GenerateSchemaFromAiAsync(interpolateSystemPromptFunc As Func(Of String, String),
                                                        llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                                        useSecondApi As Boolean,
                                                        context As ISharedContext) As Threading.Tasks.Task(Of System.Collections.Generic.List(Of ExtractionSchemaColumn))
            Dim userText = ""
            Dim systemPrompt = interpolateSystemPromptFunc(context.SP_ExtractSchema)   ' Will cause OtherPrompt etc. to be included
            Dim jsonResp = Await llmFunc(systemPrompt, userText, "", "", 0, useSecondApi, False)
            jsonResp = WebAgentInterpreter.SanitizeLlmResult(jsonResp)
            Dim schemaOnly As New System.Collections.Generic.List(Of ExtractionSchemaColumn)
            If String.IsNullOrWhiteSpace(jsonResp) Then Return schemaOnly
            Try
                Dim jt = JToken.Parse(jsonResp)
                Dim st = jt("schema")
                If st IsNot Nothing AndAlso st.Type = JTokenType.Array Then
                    For Each c In st
                        Dim name = CStr(c("name"))
                        Dim typ = CStr(c("type"))
                        If String.IsNullOrWhiteSpace(name) Then Continue For
                        schemaOnly.Add(New ExtractionSchemaColumn With {
                            .Name = name.Trim(),
                            .Type = If(String.IsNullOrWhiteSpace(typ), "text", typ.Trim().ToLowerInvariant())
                        })
                    Next
                End If
            Catch
            End Try
            Return schemaOnly
        End Function

        ''' <summary>
        ''' Appends a fixed ordered schema constraint to an existing system prompt.
        ''' </summary>
        Public Function BuildConstrainedSystemPrompt(originalInterpolatedPrompt As String,
                                                     fixedSchema As System.Collections.Generic.List(Of ExtractionSchemaColumn)) As String
            If fixedSchema Is Nothing OrElse fixedSchema.Count = 0 Then Return originalInterpolatedPrompt
            Dim sb As New StringBuilder(originalInterpolatedPrompt.Length + 300)
            sb.AppendLine(originalInterpolatedPrompt)
            sb.AppendLine()
            sb.AppendLine("FIXED ORDERED SCHEMA (DO NOT RENAME OR REORDER):")
            sb.AppendLine(String.Join(" | ", fixedSchema.Select(Function(c) c.Name)))
            sb.AppendLine("Return JSON: {""rows"":[[...]],""file_name"":""<file>""} ONLY. No schema array, no commentary.")
            Return sb.ToString()
        End Function

        ''' <summary>
        ''' Merges per-file schema and rows into the aggregate result, ensuring a <c>File</c> column and normalizing configured date columns.
        ''' </summary>
        Public Sub MergeIntoAggregate(master As FactExtractionAggregateResult,
                                      schema As System.Collections.Generic.List(Of ExtractionSchemaColumn),
                                      rows As System.Collections.Generic.List(Of ExtractionRow),
                                      sourceFileName As String,
                                      dateColumnsUser As System.Collections.Generic.List(Of Integer))

            If master.Schema.Count = 0 AndAlso schema.Count > 0 Then
                If Not schema.Any(Function(c) c.Name.Equals("File", StringComparison.OrdinalIgnoreCase)) Then
                    schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                End If
                master.Schema.AddRange(schema)
            Else
                For Each c In schema
                    If Not master.Schema.Any(Function(mc) mc.Name.Equals(c.Name, StringComparison.OrdinalIgnoreCase)) Then
                        master.Schema.Add(New ExtractionSchemaColumn With {.Name = c.Name, .Type = c.Type})
                    End If
                Next
                If Not master.Schema.Any(Function(mc) mc.Name.Equals("File", StringComparison.OrdinalIgnoreCase)) Then
                    master.Schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                End If
            End If

            Dim fileColIndex = master.Schema.FindIndex(Function(c) c.Name.Equals("File", StringComparison.OrdinalIgnoreCase))
            If fileColIndex < 0 Then
                master.Schema.Add(New ExtractionSchemaColumn With {.Name = "File", .Type = "text"})
                fileColIndex = master.Schema.Count - 1
            End If

            For Each r In rows
                Dim newRow As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                For i = 0 To master.Schema.Count - 1
                    Dim v As Object = ""
                    If i < r.Values.Count Then v = r.Values(i)
                    newRow.Values.Add(v)
                Next
                newRow.Values(fileColIndex) = sourceFileName
                For Each dc In dateColumnsUser
                    Dim idx = dc - 1
                    If idx >= 0 AndAlso idx < newRow.Values.Count Then
                        Dim raw = CStr(newRow.Values(idx))
                        If Not String.IsNullOrWhiteSpace(raw) Then
                            newRow.Values(idx) = NormalizeDate(raw)
                        End If
                    End If
                Next
                master.Rows.Add(newRow)
            Next
        End Sub

        ''' <summary>
        ''' Filters aggregate rows by keeping only rows whose configured date columns are within the optional clamp range.
        ''' </summary>
        Private Sub ApplyDateClamps(result As FactExtractionAggregateResult,
                                    dateColumnsUser As System.Collections.Generic.List(Of Integer),
                                    clampFromRaw As String,
                                    clampToRaw As String)
            If result Is Nothing OrElse result.Rows.Count = 0 OrElse dateColumnsUser Is Nothing OrElse dateColumnsUser.Count = 0 Then Return
            Dim clampFrom = ParseFlexibleDate(clampFromRaw)
            Dim clampTo = ParseFlexibleDate(clampToRaw)
            If Not clampFrom.HasValue AndAlso Not clampTo.HasValue Then Return

            Dim keep As New System.Collections.Generic.List(Of ExtractionRow)
            For Each row In result.Rows
                Dim ok As Boolean = True
                For Each dc In dateColumnsUser
                    Dim idx = dc - 1
                    If idx < 0 OrElse idx >= row.Values.Count Then Continue For
                    Dim cell = CStr(row.Values(idx))
                    Dim dt = ToComparableDate(NormalizeDate(cell))
                    If dt.HasValue Then
                        If clampFrom.HasValue AndAlso dt.Value < clampFrom.Value Then ok = False : Exit For
                        If clampTo.HasValue AndAlso dt.Value > clampTo.Value Then ok = False : Exit For
                    End If
                Next
                If ok Then keep.Add(row)
            Next
            result.Rows = keep
        End Sub

        ''' <summary>
        ''' Sorts aggregate rows by the specified 1-based column index and direction.
        ''' </summary>
        Public Sub SortAggregate(result As FactExtractionAggregateResult,
                                 sortColumn As Integer,
                                 sortDir As String)
            If result Is Nothing OrElse result.Rows.Count = 0 Then Return
            If sortColumn <= 0 Then Return
            Dim idx = sortColumn - 1
            If idx >= result.Schema.Count Then Return
            Dim typeHint = result.Schema(idx).Type
            Dim asc = Not sortDir.Equals("DESC", StringComparison.OrdinalIgnoreCase)
            result.Rows.Sort(Function(a, b)
                                 Dim av = If(idx < a.Values.Count, a.Values(idx), Nothing)
                                 Dim bv = If(idx < b.Values.Count, b.Values(idx), Nothing)
                                 Dim cmp = CompareValues(av, bv, typeHint)
                                 Return If(asc, cmp, -cmp)
                             End Function)
        End Sub

        ''' <summary>
        ''' Compares two cell values using an optional type hint.
        ''' </summary>
        Private Function CompareValues(a As Object, b As Object, typeHint As String) As Integer
            If a Is Nothing AndAlso b Is Nothing Then Return 0
            If a Is Nothing Then Return -1
            If b Is Nothing Then Return 1
            Dim sa = a.ToString()
            Dim sb = b.ToString()

            If {"date", "datetime", "time"}.Contains(typeHint) Then
                Dim da = ToComparableDate(NormalizeDate(sa))
                Dim db = ToComparableDate(NormalizeDate(sb))
                If da.HasValue AndAlso db.HasValue Then Return DateTime.Compare(da.Value, db.Value)
            End If

            Dim daNum As Double
            Dim dbNum As Double
            If Double.TryParse(sa, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, daNum) AndAlso
               Double.TryParse(sb, Globalization.NumberStyles.Any, Globalization.CultureInfo.InvariantCulture, dbNum) Then
                Return daNum.CompareTo(dbNum)
            End If

            Return String.Compare(sa, sb, StringComparison.OrdinalIgnoreCase)
        End Function

        ''' <summary>
        ''' Executes fact extraction over a set of files and returns an aggregate result with schema, rows, and errors.
        ''' </summary>
        Public Async Function RunFactExtractionAsync(filePaths As System.Collections.Generic.List(Of String),
                                             instruction As String,
                                             dateColumnsUser As System.Collections.Generic.List(Of Integer),
                                             sortColumn As Integer,
                                             sortDirection As String,
                                             doOcr As Boolean,
                                             useSecondApi As Boolean,
                                             sourceDirectory As String,
                                             interpolateSystemPromptFunc As Func(Of String, String),
                                             llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                             GetFileContentFunc As Func(Of String, Boolean, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                             context As ISharedContext,
                                             Optional fixedSchema As System.Collections.Generic.List(Of ExtractionSchemaColumn) = Nothing,
                                             Optional clampFrom As String = Nothing,
                                             Optional clampTo As String = Nothing,
                                             Optional progressCallback As Action(Of Integer, Integer, String) = Nothing,
                                             Optional mergeDateColumn As Integer = 0,
                                             Optional mergeRowsViaLlm As Boolean = False,
                                             Optional mergeInstruction As String = Nothing,
                                             Optional cancellationRequested As Func(Of Boolean) = Nothing) _
                                             As Threading.Tasks.Task(Of FactExtractionAggregateResult)

            Dim agg As New FactExtractionAggregateResult With {
                .Schema = New System.Collections.Generic.List(Of ExtractionSchemaColumn),
                .Rows = New System.Collections.Generic.List(Of ExtractionRow),
                .Errors = New System.Collections.Generic.List(Of String),
                .FailedFileNames = New System.Collections.Generic.List(Of String),
                .SourceDirectory = sourceDirectory
            }

            If filePaths Is Nothing OrElse filePaths.Count = 0 Then
                agg.Errors.Add("No input files.")
                Return agg
            End If

            For i = 0 To filePaths.Count - 1
                If cancellationRequested IsNot Nothing AndAlso cancellationRequested() Then
                    agg.Errors.Add("Cancelled by user.")
                    Exit For
                End If

                Dim path = filePaths(i)
                If progressCallback IsNot Nothing Then
                    progressCallback(i, filePaths.Count, "Processing " & System.IO.Path.GetFileName(path) &
                                         " (" & (i + 1).ToString() & " of " & filePaths.Count.ToString() & ")")
                End If

                If Not File.Exists(path) Then
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                Dim text = Await GetFileContentFunc(path, False, doOcr, False)
                If cancellationRequested IsNot Nothing AndAlso cancellationRequested() Then
                    agg.Errors.Add("Cancelled by user.")
                    Exit For
                End If
                If String.IsNullOrWhiteSpace(text) Then
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If

                Dim userText = "<TEXTTOPROCESS>" & text & "</TEXTTOPROCESS>"
                Dim sysPrompt = interpolateSystemPromptFunc(context.SP_Extract)
                If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 Then
                    sysPrompt = BuildConstrainedSystemPrompt(sysPrompt, fixedSchema)
                End If

                Dim jsonResp As String = Nothing
                Try
                    jsonResp = Await llmFunc(sysPrompt, userText, "", "", 0, useSecondApi, False)
                    jsonResp = WebAgentInterpreter.SanitizeLlmResult(jsonResp)
                Catch ex As Exception
                    agg.Errors.Add("LLM call failed for '" & System.IO.Path.GetFileName(path) & "': " & ex.Message)
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End Try
                If String.IsNullOrWhiteSpace(jsonResp) Then
                    agg.Errors.Add("Empty AI response for '" & System.IO.Path.GetFileName(path) & "'.")
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If

                Dim parsed = ParseSingleFileJson(jsonResp)
                If fixedSchema IsNot Nothing AndAlso fixedSchema.Count > 0 AndAlso parsed.schema.Count = 0 Then
                    parsed.schema.AddRange(fixedSchema.Select(Function(c) New ExtractionSchemaColumn With {.Name = c.Name, .Type = c.Type}))
                End If
                If parsed.schema.Count = 0 OrElse parsed.rows.Count = 0 Then
                    agg.Errors.Add("No rows/schema parsed for '" & System.IO.Path.GetFileName(path) & "'.")
                    agg.FailedFileNames.Add(System.IO.Path.GetFileName(path))
                    Continue For
                End If
                MergeIntoAggregate(agg, parsed.schema, parsed.rows, System.IO.Path.GetFileName(path), dateColumnsUser)
                agg.ProcessedFiles += 1
            Next

            If cancellationRequested IsNot Nothing AndAlso cancellationRequested() Then
                If progressCallback IsNot Nothing Then
                    progressCallback(agg.ProcessedFiles, Math.Max(1, filePaths.Count), "Cancelled.")
                End If
                agg.FailedFiles = agg.FailedFileNames.Count
                Return agg
            End If

            ApplyDateClamps(agg, dateColumnsUser, clampFrom, clampTo)

            ' Generic merge (any column)
            If mergeRowsViaLlm AndAlso mergeDateColumn > 0 Then
                Await MergeRowsByKeyAsync(agg,
                                          mergeDateColumn,
                                          mergeInstruction,
                                          useSecondApi,
                                          interpolateSystemPromptFunc,
                                          llmFunc,
                                          context,
                                          progressCallback,
                                          cancellationRequested)
            End If

            agg.FailedFiles = agg.FailedFileNames.Count
            If sortColumn > 0 Then SortAggregate(agg, sortColumn, sortDirection)

            If progressCallback IsNot Nothing Then
                progressCallback(filePaths.Count, filePaths.Count, If(cancellationRequested IsNot Nothing AndAlso cancellationRequested(), "Cancelled.", "Completed."))
            End If

            Return agg
        End Function


        ''' <summary>
        ''' Groups aggregate rows by a 1-based key column and merges each group into a single row (via LLM or fallback merge).
        ''' </summary>
        Private Async Function MergeRowsByKeyAsync(agg As FactExtractionAggregateResult,
                                                   keyColumn As Integer,
                                                   mergeInstruction As String,
                                                   useSecondApi As Boolean,
                                                   interpolateSystemPromptFunc As Func(Of String, String),
                                                   llmFunc As Func(Of String, String, String, String, Integer, Boolean, Boolean, Threading.Tasks.Task(Of String)),
                                                   context As ISharedContext,
                                                   progressCallback As Action(Of Integer, Integer, String),
                                                   cancellationRequested As Func(Of Boolean)) As Threading.Tasks.Task
            If agg Is Nothing OrElse agg.Rows.Count = 0 Then Return
            If keyColumn <= 0 Then Return
            Dim keyIdx = keyColumn - 1
            If keyIdx >= agg.Schema.Count Then Return

            ' Build grouping key (normalize only if date/datetime)
            Dim isDateLike = {"date", "datetime"}.Contains(agg.Schema(keyIdx).Type)
            Dim groups = agg.Rows.
                GroupBy(Function(r)
                            Dim raw = ""
                            If keyIdx < r.Values.Count AndAlso r.Values(keyIdx) IsNot Nothing Then
                                raw = r.Values(keyIdx).ToString().Trim()
                            End If
                            If String.IsNullOrWhiteSpace(raw) Then Return ""
                            If isDateLike Then Return NormalizeDate(raw)
                            Return raw
                        End Function).
                Where(Function(g) Not String.IsNullOrWhiteSpace(g.Key)).
                ToList()

            If groups.Count = 0 Then Return

            ' Reset progress bar for merge phase
            If progressCallback IsNot Nothing Then
                progressCallback(0, groups.Count, "Merging groups...")
            End If

            Dim newRows As New System.Collections.Generic.List(Of ExtractionRow)
            Dim totalGroups = groups.Count
            Dim groupIndex As Integer = 0

            For Each g In groups
                groupIndex += 1
                If cancellationRequested IsNot Nothing AndAlso cancellationRequested() Then
                    Exit For
                End If

                If progressCallback IsNot Nothing Then
                    progressCallback(groupIndex, totalGroups, $"Merging ({groupIndex}/{totalGroups}) key='{g.Key}'")
                End If

                Dim mergedRow As ExtractionRow = Nothing
                ' Build JSON input (re-use merge prompt expecting date; provide generic input)
                Dim schemaNames = agg.Schema.Select(Function(c) c.Name).ToList()
                Dim rowsArray As New System.Text.StringBuilder()
                rowsArray.Append("{""merge_key"":""" & g.Key.Replace("""", "\""") & """,""schema"":[""" & String.Join(""" ,""", schemaNames.Select(Function(s) s.Replace("""", "\"""))) & """],""rows"":[")
                Dim groupRows = g.ToList()
                For ri = 0 To groupRows.Count - 1
                    Dim r = groupRows(ri)
                    rowsArray.Append("[")
                    For ci = 0 To agg.Schema.Count - 1
                        Dim val As String = ""
                        If ci < r.Values.Count AndAlso r.Values(ci) IsNot Nothing Then val = r.Values(ci).ToString()
                        val = val.Replace("""", "\""")
                        rowsArray.Append("""" & val & """")
                        If ci < agg.Schema.Count - 1 Then rowsArray.Append(",")
                    Next
                    rowsArray.Append("]")
                    If ri < groupRows.Count - 1 Then rowsArray.Append(",")
                Next
                rowsArray.Append("]}")

                Dim systemTemplate = context.SP_MergeDateRows ' reuse existing prompt (still works)
                Dim systemPrompt = interpolateSystemPromptFunc(systemTemplate)
                Dim userText =
                    "MERGE_INSTRUCTION: " & If(String.IsNullOrWhiteSpace(mergeInstruction), "(none)", mergeInstruction) & vbCrLf &
                    "INPUT_GROUP_JSON:" & vbCrLf & rowsArray.ToString()

                Try
                    Dim resp = Await llmFunc(systemPrompt, userText, "", "", 0, useSecondApi, False)
                    resp = WebAgentInterpreter.SanitizeLlmResult(resp)
                    If Not String.IsNullOrWhiteSpace(resp) Then
                        Dim jt = Newtonsoft.Json.Linq.JToken.Parse(resp)
                        Dim valsTok = jt("values")
                        If valsTok IsNot Nothing AndAlso valsTok.Type = Newtonsoft.Json.Linq.JTokenType.Array Then
                            Dim er As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
                            For ci = 0 To agg.Schema.Count - 1
                                Dim vTok = valsTok(ci)
                                er.Values.Add(If(vTok Is Nothing, "", vTok.ToString()))
                            Next
                            mergedRow = er
                        End If
                    End If
                Catch ex As Exception
                    agg.Errors.Add("Merge LLM failed for key " & g.Key & ": " & ex.Message)
                End Try

                If mergedRow Is Nothing Then
                    mergedRow = FallbackMergeRowsGeneric(groupRows, agg.Schema, keyIdx, isDateLike)
                End If
                newRows.Add(mergedRow)
            Next

            ' Preserve rows without a usable key + merged ones
            Dim unGrouped = agg.Rows.Where(Function(r)
                                               Dim raw = ""
                                               If keyIdx < r.Values.Count AndAlso r.Values(keyIdx) IsNot Nothing Then
                                                   raw = r.Values(keyIdx).ToString().Trim()
                                               End If
                                               If String.IsNullOrWhiteSpace(raw) Then Return True
                                               If isDateLike Then
                                                   Return String.IsNullOrWhiteSpace(NormalizeDate(raw))
                                               End If
                                               Return False
                                           End Function).ToList()

            agg.Rows = unGrouped.Concat(newRows).ToList()

            ' Restore progress bar to file phase completion if not cancelled
            If progressCallback IsNot Nothing Then
                progressCallback(totalGroups, totalGroups, If(cancellationRequested IsNot Nothing AndAlso cancellationRequested(), "Merge cancelled.", "Merge completed"))
            End If
        End Function



        ''' <summary>
        ''' Fallback group merge that selects/concatenates values across rows using schema type hints.
        ''' </summary>
        Private Function FallbackMergeRowsGeneric(rows As System.Collections.Generic.List(Of ExtractionRow),
                                                  schema As System.Collections.Generic.List(Of ExtractionSchemaColumn),
                                                  keyColIdx As Integer,
                                                  keyIsDate As Boolean) As ExtractionRow
            Dim merged As New ExtractionRow With {.Values = New System.Collections.Generic.List(Of Object)()}
            For i = 0 To schema.Count - 1
                If i = keyColIdx Then
                    Dim keyVal = ""
                    If rows(0).Values.Count > keyColIdx AndAlso rows(0).Values(keyColIdx) IsNot Nothing Then
                        keyVal = rows(0).Values(keyColIdx).ToString()
                        If keyIsDate Then keyVal = NormalizeDate(keyVal)
                    End If
                    merged.Values.Add(keyVal)
                    Continue For
                End If
                Dim colType = schema(i).Type
                Dim collected As New System.Text.StringBuilder()
                Dim chosen As Object = ""
                For Each r In rows
                    If i < r.Values.Count Then
                        Dim v = r.Values(i)
                        If v IsNot Nothing Then
                            Dim s = v.ToString().Trim()
                            If s.Length > 0 Then
                                If {"number", "date", "datetime"}.Contains(colType) Then
                                    chosen = s
                                    Exit For
                                Else
                                    If collected.Length > 0 Then collected.Append(" | ")
                                    collected.Append(s)
                                    If String.IsNullOrWhiteSpace(CStr(chosen)) Then chosen = s
                                End If
                            End If
                        End If
                    End If
                Next
                If {"text", "other"}.Contains(colType) Then
                    merged.Values.Add(If(collected.Length > 0, collected.ToString(), CStr(chosen)))
                Else
                    merged.Values.Add(chosen)
                End If
            Next
            Return merged
        End Function


    End Module
End Namespace