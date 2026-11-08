' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Format-specific validation/normalization is isolated from batch orchestration.
' The model declares table intent; native workbook styles and committed content are authoritative.
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq

Namespace BatchJobs
    Friend NotInheritable Class OutputContract
        Public Property Schema As System.Int32 = 1
        Public Property Conventions As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
        Public Property Outputs As New System.Collections.Generic.List(Of OutputContractEntry)()
    End Class

    Friend NotInheritable Class OutputContractEntry
        Public Property Name As System.String = ""
        Public Property Kind As System.String = ""
        Public Property Worksheets As New System.Collections.Generic.List(Of WorksheetContract)()
    End Class

    Friend NotInheritable Class WorksheetContract
        Public Property Name As System.String = ""
        Public Property HeaderRow As System.Int32 = 1
        Public Property FirstDataRow As System.Int32 = 2
        Public Property Columns As New System.Collections.Generic.List(Of System.String)()
    End Class

    Friend NotInheritable Class OutputPolicy
        Private Shared ReadOnly S As System.Xml.Linq.XNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        Private Shared ReadOnly R As System.Xml.Linq.XNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        Private Shared ReadOnly P As System.Xml.Linq.XNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
        Private Sub New()
        End Sub

        Public Shared Function Guidance() As System.String
            Return "CUMULATIVE OUTPUT CONTRACT: Before the first successful item, create batch-output-contract.json in the workspace root with " &
                "{""Schema"":1,""Outputs"":[{""Name"":""your-output.xlsx"",""Kind"":""xlsx"",""Worksheets"":[{""Name"":""your-sheet"",""HeaderRow"":1,""FirstDataRow"":2,""Columns"":[""exact first column label"",""exact second column label""]}]}]}. " &
                "For a Markdown output use Kind=markdown and Worksheets=[]. Choose names, columns and data fields from the user's task, not this example. " &
                "Also declare Conventions as a string-to-string object for the task's fixed language, terminology, dates, units and missing-value notation. Read and reuse it on every run. Declare every cumulative output. This file is control data, NOT a user deliverable. Use text_write with explicit intermediate artifact metadata and artifact_delivery_intent=none. " &
                "Read the carried workbook/document AND the locked contract before editing. All later runs MUST use the identical contract, stable output names, sheet names, header rows and column order. Physical publication suffixes are transport names, not a reason to change this contract. " &
                "Treat the first data row as the native style exemplar. Append new data rows; preserve every prior populated cell, formula, header and native style. " &
                "Native-preserving XLSX edits retain the existing validation path. If a writer reserializes committed styles/layout, the host can preserve the committed native package and transfer only validated new row payloads. " &
                "That controlled preservation path cannot also import changed non-row features (for example charts, comments, hyperlinks or named ranges); use a native-preserving edit for those. Existing content/formula/header and artifact checks remain binding. " &
                "Do not rebuild the workbook, rename sheets, rephrase previous rows or add per-file titles/headings. Keep any totals append-safe rather than below the data. " &
                "For Markdown append to the exact prior content using its established row/section convention. Do not rewrite its prefix. " &
                "Only XLSX and Markdown have strict structural validation; for another output format the user must explicitly disable 'Validate cumulative layout'. " &
                "You are working on a COPY: a failed attempt will be discarded, and retries restart from the previous verified checkpoint. " &
                "A successful item must register and return all cumulative outputs, not only a description of changes."
        End Function

        Public Shared Function ValidateAndNormalize(rawContract As System.String, lockedContract As System.String,
                                                   candidate As System.String, previous As System.String,
                                                   report As System.Action(Of System.String)) As System.String
            If System.String.IsNullOrWhiteSpace(rawContract) OrElse rawContract.Length > 100000 Then Throw New System.IO.InvalidDataException("Missing or oversized batch-output-contract.json.")
            Dim contract = Newtonsoft.Json.JsonConvert.DeserializeObject(Of OutputContract)(rawContract)
            If contract Is Nothing OrElse contract.Schema <> 1 OrElse contract.Outputs Is Nothing OrElse contract.Outputs.Count = 0 Then Throw New System.IO.InvalidDataException("The cumulative output contract must declare its outputs.")
            If contract.Conventions Is Nothing OrElse contract.Conventions.Any(Function(pair) System.String.IsNullOrWhiteSpace(pair.Key) OrElse pair.Value Is Nothing) Then Throw New System.IO.InvalidDataException("Conventions must be a string-to-string object.")
            Dim names As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each entry In contract.Outputs
                If entry Is Nothing OrElse System.String.IsNullOrWhiteSpace(entry.Name) OrElse entry.Name <> System.IO.Path.GetFileName(entry.Name) OrElse Not names.Add(entry.Name) Then Throw New System.IO.InvalidDataException("Cumulative output names must be unique basenames.")
                entry.Kind = If(entry.Kind, "").Trim().ToLowerInvariant()
                If entry.Kind <> "xlsx" AndAlso entry.Kind <> "markdown" Then Throw New System.IO.InvalidDataException("Unsupported strict cumulative format: " & entry.Kind)
                If entry.Kind = "xlsx" AndAlso Not entry.Name.EndsWith(".xlsx", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("XLSX contracts require an .xlsx output.")
                If entry.Kind = "markdown" AndAlso Not entry.Name.EndsWith(".md", System.StringComparison.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Markdown contracts require an .md output.")
                If entry.Worksheets Is Nothing Then Throw New System.IO.InvalidDataException("Worksheets must be an array.")
                Dim sheets As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each sheet In entry.Worksheets
                    If sheet Is Nothing OrElse System.String.IsNullOrWhiteSpace(sheet.Name) OrElse Not sheets.Add(sheet.Name) OrElse
                       sheet.HeaderRow < 1 OrElse sheet.FirstDataRow <= sheet.HeaderRow OrElse sheet.FirstDataRow > 1048576 OrElse
                       sheet.Columns Is Nothing OrElse sheet.Columns.Count = 0 OrElse sheet.Columns.Count > 16384 OrElse
                       sheet.Columns.Any(Function(value) System.String.IsNullOrWhiteSpace(value)) Then Throw New System.IO.InvalidDataException("Invalid worksheet/header/data-row contract.")
                Next
                If entry.Kind = "xlsx" AndAlso entry.Worksheets.Count = 0 Then Throw New System.IO.InvalidDataException("An XLSX contract requires at least one data worksheet.")
            Next
            Dim canonical = Newtonsoft.Json.JsonConvert.SerializeObject(contract, Newtonsoft.Json.Formatting.Indented)
            If Not System.String.IsNullOrWhiteSpace(lockedContract) AndAlso Not Newtonsoft.Json.Linq.JToken.DeepEquals(
                Newtonsoft.Json.Linq.JToken.Parse(canonical), Newtonsoft.Json.Linq.JToken.Parse(lockedContract)) Then Throw New System.IO.InvalidDataException("The cumulative output contract changed. Use the existing contract unchanged.")
            Dim actual = System.IO.Directory.GetFiles(candidate).Select(Function(path) System.IO.Path.GetFileName(path)).ToList()
            If actual.Count <> names.Count OrElse actual.Any(Function(name) Not names.Contains(name)) Then Throw New System.IO.InvalidDataException("Delivered files do not match the cumulative output contract exactly.")
            For Each entry In contract.Outputs
                Dim path = Store.ContainedPath(candidate, entry.Name)
                If Not System.IO.File.Exists(path) OrElse (New System.IO.FileInfo(path)).Length = 0 Then Throw New System.IO.InvalidDataException("Missing/empty cumulative output: " & entry.Name)
                Dim oldPath = If(System.String.IsNullOrWhiteSpace(previous), Nothing, Store.ContainedPath(previous, entry.Name))
                If entry.Kind = "xlsx" Then
                    NormalizeWorkbook(path, oldPath, entry.Worksheets, report)
                Else
                    Dim text = System.IO.File.ReadAllText(path, New System.Text.UTF8Encoding(False, True)).Replace(Microsoft.VisualBasic.Constants.vbCrLf, Microsoft.VisualBasic.Constants.vbLf)
                    If oldPath IsNot Nothing Then
                        Dim oldText = System.IO.File.ReadAllText(oldPath, New System.Text.UTF8Encoding(False, True)).Replace(Microsoft.VisualBasic.Constants.vbCrLf, Microsoft.VisualBasic.Constants.vbLf)
                        If Not text.StartsWith(oldText, System.StringComparison.Ordinal) Then Throw New System.IO.InvalidDataException("The Markdown output changed previously committed content: " & entry.Name)
                    End If
                    report.Invoke("Validated cumulative Markdown prefix: " & entry.Name)
                End If
            Next
            Return canonical
        End Function

        Private NotInheritable Class Book
            Public Property SheetNames As New System.Collections.Generic.List(Of System.String)()
            Public Property Sheets As New System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XDocument)(System.StringComparer.Ordinal)
            Public Property Paths As New System.Collections.Generic.Dictionary(Of System.String, System.String)(System.StringComparer.Ordinal)
            Public Property Strings As New System.Collections.Generic.List(Of System.String)()
            Public Property StringItems As New System.Collections.Generic.List(Of System.Xml.Linq.XElement)()
            Public Property Styles As System.Xml.Linq.XDocument
        End Class

        Private Shared Function ReadPart(zip As System.IO.Compression.ZipArchive, name As System.String) As System.Xml.Linq.XDocument
            Dim entry = zip.GetEntry(name)
            If entry Is Nothing OrElse entry.Length > 67108864 Then Throw New System.IO.InvalidDataException("Missing/oversized workbook part: " & name)
            Using stream = entry.Open()
                Using reader = System.Xml.XmlReader.Create(stream, New System.Xml.XmlReaderSettings With {.DtdProcessing = System.Xml.DtdProcessing.Prohibit, .XmlResolver = Nothing, .MaxCharactersInDocument = 67108864})
                    Return System.Xml.Linq.XDocument.Load(reader)
                End Using
            End Using
        End Function

        Private Shared Function ReadBook(path As System.String) As Book
            Dim result As New Book()
            Using zip = System.IO.Compression.ZipFile.OpenRead(path)
                If zip.GetEntry("xl/sharedStrings.xml") IsNot Nothing Then
                    result.StringItems = ReadPart(zip, "xl/sharedStrings.xml").Descendants(S + "si").Select(Function(item) New System.Xml.Linq.XElement(item)).ToList()
                    result.Strings = result.StringItems.Select(Function(item) System.String.Concat(item.Descendants(S + "t").Select(Function(t) t.Value))).ToList()
                End If
                result.Styles = ReadPart(zip, "xl/styles.xml")
                Dim relations = ReadPart(zip, "xl/_rels/workbook.xml.rels")
                For Each sheet In ReadPart(zip, "xl/workbook.xml").Descendants(S + "sheet")
                    Dim id = CStr(sheet.Attribute(R + "id"))
                    Dim relation = relations.Descendants(P + "Relationship").Single(Function(item) CStr(item.Attribute("Id")) = id)
                    If CStr(relation.Attribute("TargetMode")) = "External" Then Throw New System.IO.InvalidDataException("A worksheet cannot be an external relationship.")
                    Dim target = CStr(relation.Attribute("Target"))
                    Dim uri As New System.Uri(New System.Uri("http://package/xl/workbook.xml"), target)
                    If uri.Host <> "package" Then Throw New System.IO.InvalidDataException("Invalid worksheet target.")
                    Dim name = CStr(sheet.Attribute("name"))
                    Dim part = System.Uri.UnescapeDataString(uri.AbsolutePath.TrimStart("/"c))
                    result.SheetNames.Add(name)
                    result.Sheets.Add(name, ReadPart(zip, part))
                    result.Paths.Add(name, part)
                Next
            End Using
            Return result
        End Function

        Private Shared Function Cells(sheet As System.Xml.Linq.XDocument) As System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XElement)
            Return sheet.Descendants(S + "c").ToDictionary(Function(cell) CStr(cell.Attribute("r")), Function(cell) cell, System.StringComparer.Ordinal)
        End Function

        Private Shared Function CellValue(cell As System.Xml.Linq.XElement, book As Book) As System.String
            If cell Is Nothing Then Return ""
            Dim formula = cell.Element(S + "f")
            If formula IsNot Nothing Then Return "formula:" & formula.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)
            Dim kind = CStr(cell.Attribute("t"))
            Dim value = If(cell.Element(S + "v") Is Nothing, "", cell.Element(S + "v").Value)
            If kind = "s" Then
                Dim index As System.Int32
                If Not System.Int32.TryParse(value, index) OrElse index < 0 OrElse index >= book.Strings.Count Then Throw New System.IO.InvalidDataException("Invalid shared string index.")
                Return "text:" & book.Strings(index)
            End If
            If kind = "inlineStr" Then Return "text:" & System.String.Concat(cell.Descendants(S + "t").Select(Function(t) t.Value))
            If kind = "str" Then Return "text:" & value
            Return If(kind, "n") & ":" & value
        End Function

        Private Shared Function HasCellPayload(cell As System.Xml.Linq.XElement) As System.Boolean
            Return cell.Element(S + "f") IsNot Nothing OrElse
                (cell.Element(S + "v") IsNot Nothing AndAlso cell.Element(S + "v").Value <> "") OrElse
                cell.Descendants(S + "t").Any(Function(value) value.Value <> "")
        End Function

        Private Shared Function LastDataRow(sheet As System.Xml.Linq.XDocument, firstDataRow As System.Int32) As System.Int32
            Return sheet.Descendants(S + "row").Where(Function(row) CInt(row.Attribute("r")) >= firstDataRow AndAlso
                row.Elements(S + "c").Any(Function(cell) HasCellPayload(cell))).Select(Function(row) CInt(row.Attribute("r"))).DefaultIfEmpty(firstDataRow - 1).Max()
        End Function

        Private Shared Function ColumnName(number As System.Int32) As System.String
            Dim result As System.String = ""
            While number > 0
                number -= 1
                result = System.Convert.ToChar(65 + number Mod 26) & result
                number = number \ 26
            End While
            Return result
        End Function

        Private Shared Function ColumnNumber(address As System.String) As System.Int32
            Dim value As System.Int32 = 0
            For Each ch In address
                If ch < "A"c OrElse ch > "Z"c Then Exit For
                value = value * 26 + System.Convert.ToInt32(ch) - 64
            Next
            Return value
        End Function

        Private Shared Sub PreserveStyleDefinitions(current As Book, previous As Book)
            For Each oldGroup In previous.Styles.Root.Elements()
                Dim newGroup = current.Styles.Root.Element(oldGroup.Name)
                If newGroup Is Nothing Then Throw New System.IO.InvalidDataException("A native workbook style group was removed.")
                Dim oldItems = oldGroup.Elements().ToList()
                Dim newItems = newGroup.Elements().ToList()
                If newItems.Count < oldItems.Count Then Throw New System.IO.InvalidDataException("Native workbook styles were removed.")
                For index As System.Int32 = 0 To oldItems.Count - 1
                    If Not System.Xml.Linq.XNode.DeepEquals(oldItems(index), newItems(index)) Then Throw New System.IO.InvalidDataException("An existing native workbook style was changed; reuse the carried workbook.")
                Next
            Next
        End Sub

        Private Shared Function ReadNativeTables(path As System.String, book As Book, name As System.String) As System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XDocument)
            Dim result As New System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XDocument)(System.StringComparer.Ordinal)
            Dim declarations = book.Sheets(name).Descendants(S + "tablePart").ToList()
            If declarations.Count = 0 Then Return result
            Dim sheetPart = book.Paths(name)
            Dim separator = sheetPart.LastIndexOf("/"c)
            Dim relationshipPart = sheetPart.Substring(0, separator + 1) & "_rels/" & sheetPart.Substring(separator + 1) & ".rels"
            Using zip = System.IO.Compression.ZipFile.OpenRead(path)
                Dim relationships = ReadPart(zip, relationshipPart)
                For Each declaration In declarations
                    Dim id = CStr(declaration.Attribute(R + "id"))
                    Dim relationship = relationships.Descendants(P + "Relationship").Single(Function(item) CStr(item.Attribute("Id")) = id)
                    If CStr(relationship.Attribute("TargetMode")) = "External" Then Throw New System.IO.InvalidDataException("A native table cannot be an external relationship.")
                    Dim uri As New System.Uri(New System.Uri("http://package/" & sheetPart), CStr(relationship.Attribute("Target")))
                    If uri.Host <> "package" Then Throw New System.IO.InvalidDataException("Invalid native table target.")
                    Dim part = System.Uri.UnescapeDataString(uri.AbsolutePath.TrimStart("/"c))
                    result.Add(part, ReadPart(zip, part))
                Next
            End Using
            Return result
        End Function

        ' A strict append contract separates data from rendering. A library may
        ' reserialize style defaults, shared strings, or extension metadata. Comparing
        ' those bytes rejects valid content; accepting the rewritten package loses
        ' native structure. Instead validate the proposed delta and apply that delta
        ' to a COPY of the immutable, authoritative predecessor package.
        Private Shared Sub PrepareNativeAppendCandidate(path As System.String, previousPath As System.String,
                                                        tables As System.Collections.Generic.List(Of WorksheetContract), report As System.Action(Of System.String))
            Dim proposal = ReadBook(path)
            Dim baseline = ReadBook(previousPath)
            If Not proposal.SheetNames.SequenceEqual(baseline.SheetNames, System.StringComparer.Ordinal) Then Throw New System.IO.InvalidDataException("Workbook sheet names/order changed.")
            Dim declarations = tables.ToDictionary(Function(table) table.Name, Function(table) table, System.StringComparer.Ordinal)
            Dim changed As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim additions As System.Int32 = 0
            Dim hasNewFormula As System.Boolean = False
            ValidateAppendPackageBoundary(path, previousPath, proposal, baseline)
            For Each sheetName In baseline.SheetNames
                Dim oldSheet = baseline.Sheets(sheetName)
                Dim newSheet = proposal.Sheets(sheetName)
                Dim before = Cells(oldSheet)
                Dim after = Cells(newSheet)
                For Each pair In before
                    If Not HasCellPayload(pair.Value) Then Continue For
                    Dim currentCell As System.Xml.Linq.XElement = Nothing
                    If Not after.TryGetValue(pair.Key, currentCell) OrElse CellValue(pair.Value, baseline) <> CellValue(currentCell, proposal) Then
                        Throw New System.IO.InvalidDataException("Previously committed cell changed: " & sheetName & "!" & pair.Key)
                    End If
                Next
                Dim declaration As WorksheetContract = Nothing
                declarations.TryGetValue(sheetName, declaration)
                Dim lastOldRow = If(declaration Is Nothing, System.Int32.MaxValue, LastDataRow(oldSheet, declaration.FirstDataRow))
                Dim data = oldSheet.Root.Element(S + "sheetData")
                If data Is Nothing Then Throw New System.IO.InvalidDataException("Missing native worksheet data: " & sheetName)
                Dim rows = data.Elements(S + "row").ToDictionary(Function(row) CInt(row.Attribute("r")), Function(row) row)
                For Each pair In after
                    If Not HasCellPayload(pair.Value) Then Continue For
                    Dim oldCell As System.Xml.Linq.XElement = Nothing
                    If before.TryGetValue(pair.Key, oldCell) AndAlso HasCellPayload(oldCell) Then Continue For
                    Dim rowNumber = CellRowNumber(pair.Key)
                    Dim column = ColumnNumber(pair.Key)
                    If declaration Is Nothing OrElse rowNumber <= lastOldRow OrElse rowNumber < declaration.FirstDataRow OrElse
                       column < 1 OrElse column > declaration.Columns.Count Then Throw New System.IO.InvalidDataException("Strict append output changed content outside a new declared data row: " & sheetName & "!" & pair.Key)
                    For Each merge In oldSheet.Descendants(S + "mergeCell")
                        Dim endpoints = CStr(merge.Attribute("ref")).Split(":"c)
                        If endpoints.Length <> 2 Then Throw New System.IO.InvalidDataException("Unsupported native merged range.")
                        If rowNumber >= CellRowNumber(endpoints(0)) AndAlso rowNumber <= CellRowNumber(endpoints(1)) AndAlso
                           column >= ColumnNumber(endpoints(0)) AndAlso column <= ColumnNumber(endpoints(1)) Then Throw New System.IO.InvalidDataException("An appended cell intersects a native merged range.")
                    Next
                    Dim cell = CopyAppendCell(pair.Value, proposal)
                    hasNewFormula = hasNewFormula OrElse cell.Element(S + "f") IsNot Nothing
                    Dim row As System.Xml.Linq.XElement = Nothing
                    If Not rows.TryGetValue(rowNumber, row) Then
                        row = New System.Xml.Linq.XElement(S + "row", New System.Xml.Linq.XAttribute("r", rowNumber))
                        rows.Add(rowNumber, row)
                        data.Add(row)
                    End If
                    ' An empty preformatted cell can be filled; its style is restored
                    ' by NormalizeWorkbook from the committed data-row exemplar.
                    Dim placeholder = row.Elements(S + "c").SingleOrDefault(Function(value) CStr(value.Attribute("r")) = pair.Key)
                    If placeholder IsNot Nothing Then placeholder.ReplaceWith(cell) Else row.Add(cell)
                    additions += 1
                    changed.Add(sheetName)
                Next
                If changed.Contains(sheetName) Then
                    Dim orderedRows = data.Elements(S + "row").OrderBy(Function(row) CInt(row.Attribute("r"))).ToList()
                    System.Xml.Linq.Extensions.Remove(data.Elements(S + "row"))
                    data.Add(orderedRows)
                    Dim dimension = oldSheet.Root.Element(S + "dimension")
                    If dimension IsNot Nothing Then
                        Dim references = Cells(oldSheet).Keys.ToList()
                        Dim maximumRow = references.Select(Function(address) CellRowNumber(address)).DefaultIfEmpty(1).Max()
                        Dim maximumColumn = references.Select(Function(address) ColumnNumber(address)).DefaultIfEmpty(1).Max()
                        dimension.SetAttributeValue("ref", "A1:" & ColumnName(maximumColumn) & maximumRow.ToString(System.Globalization.CultureInfo.InvariantCulture))
                    End If
                End If
            Next
            Dim prepared = path & ".native-append.pending"
            Try
                System.IO.File.Copy(previousPath, prepared, False)
                Using zip = System.IO.Compression.ZipFile.Open(prepared, System.IO.Compression.ZipArchiveMode.Update)
                    For Each sheetName In changed
                        Dim part = baseline.Paths(sheetName)
                        zip.GetEntry(part).Delete()
                        Using stream = zip.CreateEntry(part, System.IO.Compression.CompressionLevel.Optimal).Open()
                            baseline.Sheets(sheetName).Save(stream)
                        End Using
                    Next
                    ' Added values can invalidate existing formula caches too,
                    ' even when the new rows contain no formulas of their own.
                    If additions > 0 AndAlso (hasNewFormula OrElse baseline.Sheets.Values.Any(Function(sheet) sheet.Descendants(S + "f").Any())) Then
                        Dim workbook = ReadPart(zip, "xl/workbook.xml")
                        Dim calculation = workbook.Root.Element(S + "calcPr")
                        If calculation Is Nothing Then
                            calculation = New System.Xml.Linq.XElement(S + "calcPr")
                            Dim successors = New System.String() {"oleSize", "customWorkbookViews", "pivotCaches", "smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"}
                            Dim nextElement = workbook.Root.Elements().FirstOrDefault(Function(element) element.Name.Namespace = S AndAlso successors.Contains(element.Name.LocalName))
                            If nextElement Is Nothing Then workbook.Root.Add(calculation) Else nextElement.AddBeforeSelf(calculation)
                        End If
                        calculation.SetAttributeValue("fullCalcOnLoad", "1")
                        calculation.SetAttributeValue("forceFullCalc", "1")
                        zip.GetEntry("xl/workbook.xml").Delete()
                        Using stream = zip.CreateEntry("xl/workbook.xml", System.IO.Compression.CompressionLevel.Optimal).Open()
                            workbook.Save(stream)
                        End Using
                    End If
                End Using
                System.IO.File.Copy(prepared, path, True)
            Finally
                If System.IO.File.Exists(prepared) Then System.IO.File.Delete(prepared)
            End Try
            report.Invoke("Strict append: preserved the committed native package/styles/layout and transferred " & additions.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                " new-cell payload(s). Proposed style/layout reserialization was not adopted: " & System.IO.Path.GetFileName(path))
        End Sub

        Private Shared Function CellRowNumber(address As System.String) As System.Int32
            If System.String.IsNullOrWhiteSpace(address) OrElse Not System.Text.RegularExpressions.Regex.IsMatch(address, "\A[A-Z]{1,3}[1-9][0-9]{0,6}\z") Then Throw New System.IO.InvalidDataException("Invalid cell address.")
            Dim row = System.Int32.Parse(System.Text.RegularExpressions.Regex.Match(address, "[0-9]+$").Value, System.Globalization.CultureInfo.InvariantCulture)
            If row > 1048576 OrElse ColumnNumber(address) > 16384 Then Throw New System.IO.InvalidDataException("Cell address exceeds worksheet bounds.")
            Return row
        End Function

        Private Shared Function CopyAppendCell(source As System.Xml.Linq.XElement, book As Book) As System.Xml.Linq.XElement
            If source.Attributes().Any(Function(attr) attr.Name.NamespaceName <> "" OrElse Not (New System.String() {"r", "s", "t"}).Contains(attr.Name.LocalName)) OrElse
               source.Elements().Any(Function(element) element.Name <> S + "v" AndAlso element.Name <> S + "f" AndAlso element.Name <> S + "is") Then
                Throw New System.IO.InvalidDataException("An appended cell contains unsupported metadata; it was not silently discarded.")
            End If
            Dim kind = If(CStr(source.Attribute("t")), "n")
            If Not (New System.String() {"n", "b", "e", "d", "str", "inlineStr", "s"}).Contains(kind, System.StringComparer.Ordinal) OrElse
               source.Elements().GroupBy(Function(node) node.Name).Any(Function(group) group.Count() > 1) Then Throw New System.IO.InvalidDataException("Invalid appended cell payload.")
            Dim formula = source.Element(S + "f")
            If (kind = "s" OrElse kind = "inlineStr") AndAlso formula IsNot Nothing Then Throw New System.IO.InvalidDataException("Formula string results must use a scalar string cache.")
            If kind = "inlineStr" AndAlso (source.Element(S + "is") Is Nothing OrElse source.Element(S + "v") IsNot Nothing) Then Throw New System.IO.InvalidDataException("Invalid inline string payload.")
            If formula IsNot Nothing AndAlso formula.HasAttributes Then Throw New System.IO.InvalidDataException("Shared/array formula additions require a native workbook operation, not strict row append.")
            Dim cell As New System.Xml.Linq.XElement(source)
            cell.SetAttributeValue("s", Nothing)
            If CStr(cell.Attribute("t")) = "s" Then
                Dim index As System.Int32
                If cell.Element(S + "v") Is Nothing OrElse Not System.Int32.TryParse(cell.Element(S + "v").Value, index) OrElse index < 0 OrElse index >= book.StringItems.Count Then Throw New System.IO.InvalidDataException("Invalid appended shared string.")
                Dim inlineText As New System.Xml.Linq.XElement(book.StringItems(index))
                inlineText.Name = S + "is"
                cell.Element(S + "v").Remove()
                cell.SetAttributeValue("t", "inlineStr")
                cell.Add(inlineText)
            End If
            Return cell
        End Function

        Private Shared Sub ValidateAppendPackageBoundary(proposedPath As System.String, previousPath As System.String, proposed As Book, previous As Book)
            Using proposalZip = System.IO.Compression.ZipFile.OpenRead(proposedPath), baselineZip = System.IO.Compression.ZipFile.OpenRead(previousPath)
                Dim oldWorkbook = ReadPart(baselineZip, "xl/workbook.xml")
                Dim newWorkbook = ReadPart(proposalZip, "xl/workbook.xml")
                Dim oldEpoch = oldWorkbook.Root.Element(S + "workbookPr")
                Dim newEpoch = newWorkbook.Root.Element(S + "workbookPr")
                Dim old1904 = If(oldEpoch Is Nothing, "0", If(CStr(oldEpoch.Attribute("date1904")), "0"))
                Dim new1904 = If(newEpoch Is Nothing, "0", If(CStr(newEpoch.Attribute("date1904")), "0"))
                If (old1904 = "1" OrElse old1904 = "true") <> (new1904 = "1" OrElse new1904 = "true") Then Throw New System.IO.InvalidDataException("Workbook date epoch changed.")
                Dim oldNames = oldWorkbook.Descendants(S + "definedName").Select(Function(node) node.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                Dim newNames = newWorkbook.Descendants(S + "definedName").Select(Function(node) node.ToString(System.Xml.Linq.SaveOptions.DisableFormatting)).OrderBy(Function(value) value, System.StringComparer.Ordinal).ToList()
                If Not oldNames.SequenceEqual(newNames, System.StringComparer.Ordinal) Then Throw New System.IO.InvalidDataException("Defined names changed outside the row-append contract.")
                For Each sheetName In previous.SheetNames
                    ' Existing features stay in the baseline. New features are not row
                    ' payload and must not disappear unnoticed during native preservation.
                    For Each localName In New System.String() {"hyperlink", "mergeCell", "dataValidation", "conditionalFormatting", "drawing", "legacyDrawing", "oleObject", "control"}
                        Dim oldNodes = previous.Sheets(sheetName).Descendants(S + localName).ToList()
                        For Each node In proposed.Sheets(sheetName).Descendants(S + localName)
                            If Not oldNodes.Any(Function(priorNode) System.Xml.Linq.XNode.DeepEquals(priorNode, node)) Then Throw New System.IO.InvalidDataException("Unsupported added/changed non-row feature in strict append: " & sheetName & "/" & localName)
                        Next
                    Next
                    Dim oldTables = ReadNativeTables(previousPath, previous, sheetName)
                    Dim newTables = ReadNativeTables(proposedPath, proposed, sheetName)
                    If newTables.Count > oldTables.Count Then Throw New System.IO.InvalidDataException("Strict append cannot introduce new native tables.")
                    For Each table In newTables.Values
                        Dim oldTable = oldTables.Values.SingleOrDefault(Function(value) CStr(value.Root.Attribute("name")) = CStr(table.Root.Attribute("name")))
                        If oldTable Is Nothing OrElse CStr(oldTable.Root.Attribute("displayName")) <> CStr(table.Root.Attribute("displayName")) Then Throw New System.IO.InvalidDataException("A native table identity changed.")
                        Dim oldColumns = oldTable.Root.Descendants(S + "tableColumn").ToList()
                        Dim newColumns = table.Root.Descendants(S + "tableColumn").ToList()
                        If oldColumns.Count <> newColumns.Count Then Throw New System.IO.InvalidDataException("Native table columns changed.")
                        For index As System.Int32 = 0 To oldColumns.Count - 1
                            For Each attr In New System.String() {"name", "totalsRowFunction", "totalsRowLabel"}
                                If CStr(oldColumns(index).Attribute(attr)) <> CStr(newColumns(index).Attribute(attr)) Then Throw New System.IO.InvalidDataException("Native table column semantics changed.")
                            Next
                            For Each childName In New System.String() {"calculatedColumnFormula", "totalsRowFormula"}
                                If Not System.Xml.Linq.XNode.DeepEquals(oldColumns(index).Element(S + childName), newColumns(index).Element(S + childName)) Then Throw New System.IO.InvalidDataException("Native table formulas changed.")
                            Next
                        Next
                    Next
                Next
                ValidateAppendRelationships(proposalZip, baselineZip)
                Dim handled As New System.Collections.Generic.HashSet(Of System.String)(previous.Paths.Values, System.StringComparer.Ordinal)
                handled.UnionWith(proposed.Paths.Values)
                handled.UnionWith(New System.String() {"[Content_Types].xml", "xl/workbook.xml", "xl/styles.xml", "xl/sharedStrings.xml", "xl/calcChain.xml", "docProps/core.xml", "docProps/app.xml"})
                For Each entry In proposalZip.Entries
                    If entry.FullName.EndsWith("/", System.StringComparison.Ordinal) OrElse handled.Contains(entry.FullName) OrElse
                       entry.FullName.EndsWith(".rels", System.StringComparison.Ordinal) OrElse entry.FullName.StartsWith("xl/tables/", System.StringComparison.Ordinal) Then Continue For
                    Dim oldEntry = baselineZip.GetEntry(entry.FullName)
                    If oldEntry Is Nothing OrElse EntryHash(entry) <> EntryHash(oldEntry) Then Throw New System.IO.InvalidDataException("Strict append cannot add/change a non-row package part: " & entry.FullName)
                Next
            End Using
        End Sub

        Private Shared Sub ValidateAppendRelationships(proposal As System.IO.Compression.ZipArchive, baseline As System.IO.Compression.ZipArchive)
            For Each entry In proposal.Entries.Where(Function(value) value.FullName.EndsWith(".rels", System.StringComparison.Ordinal))
                Dim oldEntry = baseline.GetEntry(entry.FullName)
                Dim oldRelations = If(oldEntry Is Nothing, New System.Collections.Generic.List(Of System.Xml.Linq.XElement)(), ReadPart(baseline, entry.FullName).Root.Elements(P + "Relationship").ToList())
                For Each relation In ReadPart(proposal, entry.FullName).Root.Elements(P + "Relationship")
                    Dim kind = CStr(relation.Attribute("Type"))
                    ' New string tables are consumed as inline payload; calculation
                    ' caches are not imported. Neither grants access to external parts.
                    If (kind = R.NamespaceName & "/sharedStrings" OrElse kind = R.NamespaceName & "/calcChain") AndAlso CStr(relation.Attribute("TargetMode")) <> "External" Then Continue For
                    Dim target = RelationshipTarget(entry.FullName, relation)
                    If Not oldRelations.Any(Function(oldRelation) CStr(oldRelation.Attribute("Type")) = kind AndAlso
                        CStr(oldRelation.Attribute("TargetMode")) = CStr(relation.Attribute("TargetMode")) AndAlso
                        RelationshipTarget(entry.FullName, oldRelation) = target) Then Throw New System.IO.InvalidDataException("Strict append cannot introduce or redirect a package relationship: " & entry.FullName)
                Next
            Next
        End Sub

        Private Shared Function RelationshipTarget(relationshipPart As System.String, relation As System.Xml.Linq.XElement) As System.String
            Dim target = CStr(relation.Attribute("Target"))
            If System.String.IsNullOrWhiteSpace(target) Then Throw New System.IO.InvalidDataException("Missing package relationship target.")
            If CStr(relation.Attribute("TargetMode")) = "External" Then Return target
            Dim owner As System.String = "/"
            If relationshipPart <> "_rels/.rels" Then
                Dim separator = relationshipPart.LastIndexOf("/_rels/", System.StringComparison.Ordinal)
                If separator < 0 Then Throw New System.IO.InvalidDataException("Invalid package relationship part.")
                owner = "/" & relationshipPart.Substring(0, separator + 1) & relationshipPart.Substring(separator + 7)
                owner = owner.Substring(0, owner.Length - 5)
            End If
            Dim resolved As New System.Uri(New System.Uri("http://package" & owner), target)
            If resolved.Host <> "package" Then Throw New System.IO.InvalidDataException("Invalid internal package relationship.")
            Return resolved.AbsolutePath
        End Function

        Private Shared Function EntryHash(entry As System.IO.Compression.ZipArchiveEntry) As System.String
            Using stream = entry.Open(), sha = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(sha.ComputeHash(stream))
            End Using
        End Function

        ' Preserve the existing accepted native-update path. Rebase only when the
        ' old validator would reject committed presentation/serialization changes.
        ' Content, output identity, schema and feature checks are never bypassed.
        Private Shared Function RequiresNativeAppendPreservation(path As System.String, previousPath As System.String,
                                                                 tables As System.Collections.Generic.List(Of WorksheetContract)) As System.Boolean
            Dim current = ReadBook(path)
            Dim previous = ReadBook(previousPath)
            If Not current.SheetNames.SequenceEqual(previous.SheetNames, System.StringComparer.Ordinal) Then Return False
            Try
                PreserveStyleDefinitions(current, previous)
            Catch ex As System.IO.InvalidDataException
                ' This pure helper checks presentation only; no content or files
                ' have been mutated. The preservation gate separately proves data.
                Return True
            End Try
            For Each name In previous.Sheets.Keys
                Dim before = Cells(previous.Sheets(name))
                Dim after = Cells(current.Sheets(name))
                For Each nodeName In New System.String() {"cols", "sheetFormatPr", "mergeCells", "pageMargins", "pageSetup", "printOptions", "headerFooter"}
                    If Not System.Xml.Linq.XNode.DeepEquals(previous.Sheets(name).Root.Element(S + nodeName), current.Sheets(name).Root.Element(S + nodeName)) Then Return True
                Next
                For Each pair In before
                    If Not HasCellPayload(pair.Value) Then Continue For
                    Dim cell As System.Xml.Linq.XElement = Nothing
                    If after.TryGetValue(pair.Key, cell) AndAlso CStr(cell.Attribute("s")) <> CStr(pair.Value.Attribute("s")) Then Return True
                Next
            Next
            For Each table In tables
                If Not previous.Sheets.ContainsKey(table.Name) Then Return False
                Dim lastOldRow = LastDataRow(previous.Sheets(table.Name), table.FirstDataRow)
                Dim currentRows = current.Sheets(table.Name).Descendants(S + "row").ToDictionary(Function(row) CInt(row.Attribute("r")), Function(row) row)
                For Each oldRow In previous.Sheets(table.Name).Descendants(S + "row").Where(Function(row) CInt(row.Attribute("r")) >= table.FirstDataRow AndAlso CInt(row.Attribute("r")) <= lastOldRow)
                    Dim currentRow As System.Xml.Linq.XElement = Nothing
                    If Not currentRows.TryGetValue(CInt(oldRow.Attribute("r")), currentRow) Then Continue For
                    For Each attr In New System.String() {"ht", "customHeight", "s", "customFormat"}
                        If CStr(oldRow.Attribute(attr)) <> CStr(currentRow.Attribute(attr)) Then Return True
                    Next
                Next
                Dim beforeTables = ReadNativeTables(previousPath, previous, table.Name)
                Dim afterTables = ReadNativeTables(path, current, table.Name)
                If Not (New System.Collections.Generic.HashSet(Of System.String)(beforeTables.Keys)).SetEquals(afterTables.Keys) Then Return False
                For Each part In beforeTables.Keys
                    For Each nodeName In New System.String() {"tableColumns", "tableStyleInfo"}
                        If Not System.Xml.Linq.XNode.DeepEquals(beforeTables(part).Root.Element(S + nodeName), afterTables(part).Root.Element(S + nodeName)) Then Return True
                    Next
                Next
            Next
            Return False
        End Function

        Private Shared Sub NormalizeWorkbook(path As System.String, previousPath As System.String, tables As System.Collections.Generic.List(Of WorksheetContract), report As System.Action(Of System.String))
            If Not System.String.IsNullOrWhiteSpace(previousPath) AndAlso RequiresNativeAppendPreservation(path, previousPath, tables) Then
                report.Invoke("Native presentation changed; checking the bounded row-delta preservation path: " & System.IO.Path.GetFileName(path))
                PrepareNativeAppendCandidate(path, previousPath, tables, report)
            End If
            Dim current = ReadBook(path)
            Dim previous As Book = If(previousPath Is Nothing, Nothing, ReadBook(previousPath))
            If previous IsNot Nothing Then
                If Not current.SheetNames.SequenceEqual(previous.SheetNames, System.StringComparer.Ordinal) Then Throw New System.IO.InvalidDataException("Workbook sheet names/order changed.")
                PreserveStyleDefinitions(current, previous)
                For Each name In previous.Sheets.Keys
                    Dim before = Cells(previous.Sheets(name))
                    Dim after = Cells(current.Sheets(name))
                    For Each nodeName In New System.String() {"cols", "sheetFormatPr", "mergeCells", "pageMargins", "pageSetup", "printOptions", "headerFooter"}
                        If Not System.Xml.Linq.XNode.DeepEquals(previous.Sheets(name).Root.Element(S + nodeName), current.Sheets(name).Root.Element(S + nodeName)) Then
                            Throw New System.IO.InvalidDataException("Previously committed worksheet layout changed: " & name & "/" & nodeName)
                        End If
                    Next
                    For Each pair In before
                        ' Preformatted empty cells are intended for subsequent rows; only
                        ' previously populated cells are immutable content.
                        If Not HasCellPayload(pair.Value) Then Continue For
                        Dim cell As System.Xml.Linq.XElement = Nothing
                        If Not after.TryGetValue(pair.Key, cell) OrElse CellValue(cell, current) <> CellValue(pair.Value, previous) OrElse
                           CStr(cell.Attribute("s")) <> CStr(pair.Value.Attribute("s")) Then Throw New System.IO.InvalidDataException("Previously committed cell changed: " & name & "!" & pair.Key)
                    Next
                Next
            End If
            Dim normalized As System.Int32 = 0
            For Each table In tables
                If Not current.Sheets.ContainsKey(table.Name) Then Throw New System.IO.InvalidDataException("Missing worksheet: " & table.Name)
                Dim sheet = current.Sheets(table.Name)
                Dim cellsByAddress = Cells(sheet)
                For index As System.Int32 = 1 To table.Columns.Count
                    Dim address = ColumnName(index) & table.HeaderRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    Dim cell As System.Xml.Linq.XElement = Nothing
                    If Not cellsByAddress.TryGetValue(address, cell) OrElse CellValue(cell, current) <> "text:" & table.Columns(index - 1) Then Throw New System.IO.InvalidDataException("Column header mismatch: " & table.Name & "!" & address)
                Next
                Dim header = sheet.Descendants(S + "row").SingleOrDefault(Function(row) CInt(row.Attribute("r")) = table.HeaderRow)
                If header Is Nothing OrElse header.Elements(S + "c").Any(Function(cell) HasCellPayload(cell) AndAlso ColumnNumber(CStr(cell.Attribute("r"))) > table.Columns.Count) Then Throw New System.IO.InvalidDataException("The header has undeclared columns: " & table.Name)
                Dim exemplarSheet = If(previous Is Nothing, sheet, previous.Sheets(table.Name))
                Dim exemplar = exemplarSheet.Descendants(S + "row").SingleOrDefault(Function(row) CInt(row.Attribute("r")) = table.FirstDataRow)
                If exemplar Is Nothing Then Throw New System.IO.InvalidDataException("Missing native data-row exemplar in " & table.Name)
                Dim templateCells = exemplar.Elements(S + "c").ToDictionary(Function(cell) ColumnNumber(CStr(cell.Attribute("r"))), Function(cell) cell)
                Dim lastOldRow As System.Int32 = If(previous Is Nothing, table.FirstDataRow - 1, LastDataRow(previous.Sheets(table.Name), table.FirstDataRow))
                If previous IsNot Nothing Then
                    For Each oldRow In previous.Sheets(table.Name).Descendants(S + "row").Where(Function(row) CInt(row.Attribute("r")) >= table.FirstDataRow AndAlso CInt(row.Attribute("r")) <= lastOldRow)
                        Dim currentRow = sheet.Descendants(S + "row").SingleOrDefault(Function(row) CInt(row.Attribute("r")) = CInt(oldRow.Attribute("r")))
                        If currentRow Is Nothing Then Throw New System.IO.InvalidDataException("A committed data row was removed.")
                        For Each attr In New System.String() {"ht", "customHeight", "s", "customFormat"}
                            If CStr(oldRow.Attribute(attr)) <> CStr(currentRow.Attribute(attr)) Then Throw New System.IO.InvalidDataException("A committed row format changed.")
                        Next
                    Next
                End If
                For Each row In sheet.Descendants(S + "row").Where(Function(value) CInt(value.Attribute("r")) >= table.FirstDataRow AndAlso CInt(value.Attribute("r")) > lastOldRow AndAlso value.Elements(S + "c").Any(Function(cell) HasCellPayload(cell))).ToList()
                    If row.Elements(S + "c").Any(Function(cell) ColumnNumber(CStr(cell.Attribute("r"))) > table.Columns.Count) Then Throw New System.IO.InvalidDataException("A cumulative data row has extra columns.")
                    For Each attr In New System.String() {"ht", "customHeight", "s", "customFormat"}
                        row.SetAttributeValue(attr, CStr(exemplar.Attribute(attr)))
                    Next
                    Dim rowCells = row.Elements(S + "c").ToDictionary(Function(cell) ColumnNumber(CStr(cell.Attribute("r"))), Function(cell) cell)
                    For index As System.Int32 = 1 To table.Columns.Count
                        Dim cell As System.Xml.Linq.XElement = Nothing
                        If Not rowCells.TryGetValue(index, cell) Then
                            cell = New System.Xml.Linq.XElement(S + "c", New System.Xml.Linq.XAttribute("r", ColumnName(index) & CStr(row.Attribute("r"))))
                            row.Add(cell)
                        End If
                        Dim template As System.Xml.Linq.XElement = Nothing
                        templateCells.TryGetValue(index, template)
                        Dim style = If(template Is Nothing, Nothing, CStr(template.Attribute("s")))
                        If CStr(cell.Attribute("s")) <> style Then normalized += 1
                        cell.SetAttributeValue("s", style)
                    Next
                    Dim ordered = row.Elements(S + "c").OrderBy(Function(cell) ColumnNumber(CStr(cell.Attribute("r")))).ToList()
                    System.Xml.Linq.Extensions.Remove(row.Elements(S + "c"))
                    row.Add(ordered)
                Next
            Next
            ' Preserve and extend existing native table ranges rather than rebuilding
            ' tables or leaving appended rows outside the original table style/filter.
            Dim tableParts As New System.Collections.Generic.Dictionary(Of System.String, System.Xml.Linq.XDocument)(System.StringComparer.Ordinal)
            For Each table In tables
                Dim nativeTables = ReadNativeTables(path, current, table.Name)
                Dim oldTables = If(previous Is Nothing, Nothing, ReadNativeTables(previousPath, previous, table.Name))
                If oldTables IsNot Nothing AndAlso Not (New System.Collections.Generic.HashSet(Of System.String)(nativeTables.Keys)).SetEquals(oldTables.Keys) Then Throw New System.IO.InvalidDataException("Native table identities changed.")
                For Each pair In nativeTables
                    Dim root = pair.Value.Root
                    If oldTables IsNot Nothing Then
                        For Each attr In New System.String() {"name", "displayName", "headerRowCount", "totalsRowCount", "totalsRowShown"}
                            If CStr(root.Attribute(attr)) <> CStr(oldTables(pair.Key).Root.Attribute(attr)) Then Throw New System.IO.InvalidDataException("A native table definition changed.")
                        Next
                        For Each partName In New System.String() {"tableColumns", "tableStyleInfo"}
                            If Not System.Xml.Linq.XNode.DeepEquals(root.Element(S + partName), oldTables(pair.Key).Root.Element(S + partName)) Then Throw New System.IO.InvalidDataException("Native table columns/styles changed.")
                        Next
                    End If
                    Dim range = CStr(root.Attribute("ref"))
                    Dim parts = If(range, "").Split(":"c)
                    Dim start = "A" & table.HeaderRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    If parts.Length <> 2 OrElse parts(0) <> start OrElse ColumnNumber(parts(1)) <> table.Columns.Count Then Throw New System.IO.InvalidDataException("Native table range does not match the declared cumulative columns.")
                    If CStr(root.Attribute("totalsRowCount")) IsNot Nothing AndAlso CStr(root.Attribute("totalsRowCount")) <> "0" Then Throw New System.IO.InvalidDataException("Strict cumulative tables cannot have a trailing totals row; keep summaries on a separate sheet.")
                    Dim headers = root.Descendants(S + "tableColumn").Select(Function(column) CStr(column.Attribute("name"))).ToList()
                    If Not headers.SequenceEqual(table.Columns, System.StringComparer.Ordinal) Then Throw New System.IO.InvalidDataException("Native table headers do not match the cumulative contract.")
                    Dim endRow = LastDataRow(current.Sheets(table.Name), table.FirstDataRow)
                    Dim newRange = start & ":" & ColumnName(table.Columns.Count) & endRow.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    root.SetAttributeValue("ref", newRange)
                    Dim filter = root.Element(S + "autoFilter")
                    If filter IsNot Nothing Then filter.SetAttributeValue("ref", newRange)
                    tableParts.Add(pair.Key, pair.Value)
                    If range <> newRange Then report.Invoke("Extended existing native table range to " & newRange & ": " & table.Name)
                Next
            Next
            Using zip = System.IO.Compression.ZipFile.Open(path, System.IO.Compression.ZipArchiveMode.Update)
                For Each pair In tableParts
                    zip.GetEntry(pair.Key).Delete()
                    Using stream = zip.CreateEntry(pair.Key, System.IO.Compression.CompressionLevel.Optimal).Open()
                        pair.Value.Save(stream)
                    End Using
                Next
                For Each table In tables
                    Dim part = current.Paths(table.Name)
                    zip.GetEntry(part).Delete()
                    Using stream = zip.CreateEntry(part, System.IO.Compression.CompressionLevel.Optimal).Open()
                        current.Sheets(table.Name).Save(stream)
                    End Using
                Next
            End Using
            report.Invoke("Validated workbook content/headers; normalized " & normalized.ToString(System.Globalization.CultureInfo.InvariantCulture) & " new-cell styles: " & System.IO.Path.GetFileName(path))
        End Sub
    End Class
End Namespace
