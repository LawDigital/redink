' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Admin.Model.vb
' Purpose: Typed admin editing; runtime contracts remain authoritative.
'
' Architecture:
'  - Part of the AutoPilot Scan Handler; configuration and recipient authority
'    remain in the validated, persisted scan register.
'  - Document/model content never grants transport or filesystem permissions.
'  - Host lifecycle, shared UI and retained operation checkpoints are reused.
' =============================================================================
Option Explicit On
Option Strict On
Option Infer On
Imports System.Linq

Namespace ScanInbox
    Public NotInheritable Class AdminSession
        Public ReadOnly Property Path As System.String
        Public Property Configuration As Configuration
            Get
                Return _configuration
            End Get
            Private Set(value As Configuration)
                _configuration = value
            End Set
        End Property
        Private _configuration As Configuration
        Private _revision As System.Int64
        Private _hash As System.String = ""
        Private _saved As System.String = ""
        Private _replacementPending As System.Boolean
        Public Property LoadError As System.String
            Get
                Return _loadError
            End Get
            Private Set(value As System.String)
                _loadError = value
            End Set
        End Property
        Private _loadError As System.String = ""
        Public Property ReviewingBackup As System.Boolean
        Public ReadOnly Property HasChanges As System.Boolean
            Get
                Return Configuration IsNot Nothing AndAlso (_replacementPending OrElse ReviewingBackup OrElse Serialize(Configuration) <> _saved)
            End Get
        End Property

        Public Sub New(path As System.String)
            Me.Path = path
        End Sub

        Public Shared Function Copy(Of T)(value As T) As T
            Return Registry.Parse(Of T)(Serialize(value))
        End Function

        Private Shared Function Serialize(value As System.Object) As System.String
            Return Newtonsoft.Json.JsonConvert.SerializeObject(value)
        End Function

        Public Sub Reload()
            Configuration = Nothing
            ReviewingBackup = False
            _replacementPending = False
            LoadError = ""
            Dim exists = System.IO.File.Exists(Path)
            Dim json = If(exists, System.IO.File.ReadAllText(Path, System.Text.Encoding.UTF8), "")
            _hash = If(exists, Registry.Hash(json), "")
            _revision = 0
            Try
                Dim loaded = If(exists, Registry.Parse(Of Configuration)(json), New Configuration())
                _revision = loaded.Revision
                Registry.Validate(loaded)
                Configuration = loaded
                _saved = Serialize(loaded)
            Catch ex As System.Exception
                LoadError = ex.Message
            End Try
        End Sub

        Public Sub StartNew()
            Configuration = New Configuration With {.Revision = _revision}
            ReviewingBackup = False
            _replacementPending = True
        End Sub

        Public Sub LoadBackup()
            ' Preserve the originally observed disk hash/revision: backup is a draft, not a write.
            Dim backup = Registry.Load(Path & ".bak")
            Configuration = Copy(backup)
            Configuration.Revision = System.Math.Max(_revision, backup.Revision)
            ReviewingBackup = True
        End Sub

        Public Sub Save()
            If Configuration Is Nothing Then Throw New System.InvalidOperationException("Load a valid configuration or restore a backup first.")
            ' Validate a copy: failed validation or concurrency must not mutate the user's draft.
            Dim candidate = Copy(Configuration)
            Registry.Save(Path, candidate, _revision, _hash, expectedMissing:=(_hash = ""))
            Configuration = candidate
            _revision = candidate.Revision
            _hash = Registry.Hash(Newtonsoft.Json.JsonConvert.SerializeObject(candidate, Newtonsoft.Json.Formatting.Indented))
            _saved = Serialize(candidate)
            ReviewingBackup = False
            _replacementPending = False
            LoadError = ""
        End Sub

        Public Sub SaveChange(change As System.Action)
            Dim previous = If(Configuration Is Nothing, Nothing, Copy(Configuration))
            Dim pending = _replacementPending
            Dim backup = ReviewingBackup
            Try
                change()
                Save()
            Catch
                Configuration = previous
                _replacementPending = pending
                ReviewingBackup = backup
                Throw
            End Try
        End Sub

        Public Sub ImportFile(fileName As System.String)
            ' Parse and validate everything before replacing the current configuration.
            Dim imported = Registry.Load(fileName)
            SaveChange(Sub()
                           Configuration = Copy(imported)
                           Configuration.Revision = _revision
                       End Sub)
        End Sub

        Public Sub ExportFile(fileName As System.String)
            If Configuration Is Nothing Then Throw New System.InvalidOperationException("No valid configuration to export.")
            If System.String.Equals(System.IO.Path.GetFullPath(fileName), System.IO.Path.GetFullPath(Path), System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(System.IO.Path.GetFullPath(fileName), System.IO.Path.GetFullPath(Path & ".bak"), System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.InvalidOperationException("Choose a separate export file, not the active configuration or its automatic backup.")
            End If
            Dim snapshot = Copy(Configuration)
            Registry.Validate(snapshot)
            Registry.AtomicWrite(fileName, Newtonsoft.Json.JsonConvert.SerializeObject(snapshot, Newtonsoft.Json.Formatting.Indented))
        End Sub

        Public Shared ReadOnly ImportColumns As System.String() = {"Name", "Email", "Organization", "Active", "Aliases", "Location", "Attributes", "Instruction", "MailLanguage", "TranslateAttachments", "TranslationLanguage", "DraftReply", "ReplyOnlyWhenAppropriate", "WebResearch", "TaskInstructions", "CopySmtpAddresses"}

        Private Shared Function CsvCell(value As System.String) As System.String
            Return """" & If(value, "").Replace("""", """""") & """"
        End Function

        Private Shared Function ImportList(value As System.String) As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            Dim part As New System.Text.StringBuilder()
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim ch = value(index)
                If ch = "\"c AndAlso index + 1 < value.Length AndAlso (value(index + 1) = "\"c OrElse value(index + 1) = "|"c) Then
                    index += 1
                    part.Append(value(index))
                ElseIf ch = "|"c Then
                    If part.Length > 0 Then result.Add(part.ToString())
                    part.Clear()
                Else
                    part.Append(ch)
                End If
                index += 1
            End While
            If part.Length > 0 Then result.Add(part.ToString())
            Return result
        End Function

        Private Shared Function ExportList(values As System.Collections.Generic.List(Of System.String)) As System.String
            Return System.String.Join("|", values.Select(Function(v) v.Replace("\", "\\").Replace("|", "\|")))
        End Function

        Public Sub ExportRecipientsFile(fileName As System.String)
            If Configuration Is Nothing Then Throw New System.InvalidOperationException("No valid configuration to export.")
            If Not System.IO.Path.GetExtension(fileName).Equals(".csv", System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(System.IO.Path.GetFullPath(fileName), System.IO.Path.GetFullPath(Path), System.StringComparison.OrdinalIgnoreCase) OrElse
                System.String.Equals(System.IO.Path.GetFullPath(fileName), System.IO.Path.GetFullPath(Path & ".bak"), System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.InvalidOperationException("Choose a separate .csv export file.")
            End If
            Dim snapshot = Copy(Configuration)
            Registry.Validate(snapshot)
            Dim csv As New System.Text.StringBuilder()
            csv.AppendLine(System.String.Join(";", ImportColumns.Select(Function(c) CsvCell(c))))
            For Each person In snapshot.Recipients
                Dim values As System.String() = {person.Name, person.SmtpAddress, person.Organization, person.Active.ToString(), ExportList(person.Aliases), person.Location, ExportList(person.Attributes), person.Instruction, person.MailLanguage, person.TranslateAttachments.ToString(), person.TranslationLanguage, person.DraftReply.ToString(), person.ReplyOnlyWhenAppropriate.ToString(), person.WebResearch.ToString(), person.TaskInstructions, ExportList(person.CopySmtpAddresses)}
                csv.AppendLine(System.String.Join(";", values.Select(Function(v) CsvCell(v))))
            Next
            ' BOM enables direct opening in Excel; the importer also accepts UTF-8 without BOM.
            Registry.AtomicWrite(fileName, ChrW(&HFEFF) & csv.ToString())
        End Sub

        Private Shared Function ImportBoolean(value As System.String, column As System.String) As System.Boolean
            Select Case value.Trim().ToLowerInvariant()
                Case "true", "yes", "1" : Return True
                Case "false", "no", "0" : Return False
                Case Else : Throw New System.IO.InvalidDataException(column & ": use true/false, yes/no or 1/0.")
            End Select
        End Function

        Private Shared Function ImportPerson(fields As System.String(), columns As System.String(), existing As RegisteredRecipient) As RegisteredRecipient
            If fields.Length <> columns.Length Then Throw New System.IO.InvalidDataException("The number of values must match the column headings.")
            Dim person = If(existing Is Nothing, New RegisteredRecipient(), Copy(existing))
            For index As System.Int32 = 0 To columns.Length - 1
                Dim value = fields(index).Trim()
                ' Empty switches preserve defaults/current values; empty text/list cells clear that field.
                If value = "" AndAlso New System.String() {"active", "translateattachments", "draftreply", "replyonlywhenappropriate", "webresearch"}.Contains(columns(index).ToLowerInvariant()) Then Continue For
                Select Case columns(index).ToLowerInvariant()
                    Case "name" : person.Name = value
                    Case "email" : person.SmtpAddress = value
                    Case "copysmtpaddresses" : person.CopySmtpAddresses = ImportList(fields(index))
                    Case "organization" : person.Organization = value
                    Case "active" : person.Active = ImportBoolean(value, columns(index))
                    Case "aliases" : person.Aliases = ImportList(fields(index))
                    Case "location" : person.Location = value
                    Case "attributes" : person.Attributes = ImportList(fields(index))
                    Case "instruction" : person.Instruction = fields(index)
                    Case "maillanguage" : person.MailLanguage = value
                    Case "translateattachments" : person.TranslateAttachments = ImportBoolean(value, columns(index))
                    Case "translationlanguage" : person.TranslationLanguage = value
                    Case "draftreply" : person.DraftReply = ImportBoolean(value, columns(index))
                    Case "replyonlywhenappropriate" : person.ReplyOnlyWhenAppropriate = ImportBoolean(value, columns(index))
                    Case "webresearch" : person.WebResearch = ImportBoolean(value, columns(index))
                    Case "taskinstructions" : person.TaskInstructions = fields(index)
                End Select
            Next
            If person.Name = "" OrElse person.Name.Any(Function(c) System.Char.IsControl(c)) OrElse person.Organization.Any(Function(c) System.Char.IsControl(c)) OrElse person.Location.Any(Function(c) System.Char.IsControl(c)) Then Throw New System.IO.InvalidDataException("Enter a name; name, organization and location must each use one line.")
            person.SmtpAddress = Registry.Address(person.SmtpAddress)
            Registry.ValidatePersonalSettings(person)
            Registry.ValidateCopyRecipients(person)
            Return person
        End Function

        Public Function PreviewRecipients(text As System.String, delimiter As System.String, hasHeader As System.Boolean, Optional updateExisting As System.Boolean = False) As System.Collections.Generic.List(Of AdminImportRow)
            If Configuration Is Nothing Then Throw New System.InvalidOperationException("Load a valid configuration first.")
            If System.String.IsNullOrWhiteSpace(text) Then Throw New System.IO.InvalidDataException("Paste names and email addresses or open a CSV file.")
            If text.Length > 2000000 Then Throw New System.IO.InvalidDataException("The import is too large (maximum 2 million characters).")
            If Not New System.String() {vbTab, ";", ","}.Contains(delimiter) Then Throw New System.IO.InvalidDataException("Choose a column separator.")
            Dim rows As New System.Collections.Generic.List(Of AdminImportRow)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            Using input As New System.IO.StringReader(text.TrimStart(ChrW(&HFEFF)))
                Using parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(input)
                    parser.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
                    parser.SetDelimiters(delimiter)
                    parser.HasFieldsEnclosedInQuotes = True
                    parser.TrimWhiteSpace = False
                    Dim columns As System.String() = Nothing
                    If hasHeader AndAlso Not parser.EndOfData Then
                        columns = parser.ReadFields().Select(Function(c) c.Trim()).ToArray()
                        If columns.Distinct(System.StringComparer.OrdinalIgnoreCase).Count() <> columns.Length OrElse
                            columns.Any(Function(c) Not ImportColumns.Contains(c, System.StringComparer.OrdinalIgnoreCase)) Then Throw New System.IO.InvalidDataException("Unknown or duplicate column heading. Supported columns: " & System.String.Join(", ", ImportColumns))
                        If Not columns.Contains("Name", System.StringComparer.OrdinalIgnoreCase) OrElse Not columns.Contains("Email", System.StringComparer.OrdinalIgnoreCase) Then Throw New System.IO.InvalidDataException("Name and Email columns are required.")
                    End If
                    While Not parser.EndOfData
                        If rows.Count >= 1000 Then Throw New System.IO.InvalidDataException("Import at most 1,000 rows at a time.")
                        Dim row As New AdminImportRow With {.Line = parser.LineNumber}
                        Try
                            Dim fields = parser.ReadFields()
                            If fields Is Nothing Then Exit While
                            Dim mapped = columns
                            If Not hasHeader Then
                                If fields.Length < 2 OrElse fields.Length > 3 Then Throw New System.IO.InvalidDataException("Without headings use Name, Email and optionally Organization. Use headings for all other attributes.")
                                mapped = ImportColumns.Take(fields.Length).ToArray()
                            End If
                            Dim emailIndex = System.Array.FindIndex(mapped, Function(c) c.Equals("Email", System.StringComparison.OrdinalIgnoreCase))
                            If fields.Length <> mapped.Length Then Throw New System.IO.InvalidDataException("The number of values must match the column headings.")
                            Dim email = Registry.Address(fields(emailIndex))
                            Dim existing = Configuration.Recipients.SingleOrDefault(Function(p) p.SmtpAddress.Equals(email, System.StringComparison.OrdinalIgnoreCase))
                            row.Person = ImportPerson(fields, mapped, If(updateExisting, existing, Nothing))
                            row.Name = row.Person.Name
                            row.Email = row.Person.SmtpAddress
                            row.Organization = row.Person.Organization
                            If Configuration.RuleSets.Any(Function(r) r.ScannerSenders.Any(Function(a) a.Equals(row.Email, System.StringComparison.OrdinalIgnoreCase) OrElse row.Person.CopySmtpAddresses.Contains(a, System.StringComparer.OrdinalIgnoreCase))) Then Throw New System.IO.InvalidDataException("A scanner sender cannot be a recipient or copy recipient.")
                            If Not seen.Add(row.Email) Then
                                row.Status = "Duplicate in import — skipped"
                            ElseIf existing IsNot Nothing Then
                                If updateExisting Then
                                    If Not row.Person.Active AndAlso RecipientReferences(existing.Id) <> "" Then Throw New System.IO.InvalidDataException("Remove this recipient from recipient sets before disabling: " & RecipientReferences(existing.Id))
                                    row.Status = "Update existing recipient"
                                    row.Update = True
                                Else
                                    row.Status = "Already registered — unchanged"
                                End If
                            Else
                                row.Status = "New recipient"
                                row.Add = True
                            End If
                        Catch ex As System.Exception
                            row.Issue = ex.Message
                            row.Status = "Correct this row"
                        End Try
                        rows.Add(row)
                    End While
                End Using
            End Using
            Return rows
        End Function

        Public Function ImportRecipients(text As System.String, delimiter As System.String, hasHeader As System.Boolean, setId As System.String, Optional updateExisting As System.Boolean = False) As System.Int32
            Dim rows = PreviewRecipients(text, delimiter, hasHeader, updateExisting)
            If rows.Any(Function(r) r.Issue <> "") Then Throw New System.IO.InvalidDataException("Correct the marked rows and preview again. Nothing has been imported.")
            Dim additions = rows.Where(Function(r) r.Add OrElse r.Update).ToList()
            If additions.Count = 0 Then Throw New System.IO.InvalidDataException("No recipients to add or update.")
            SaveChange(Sub()
                           Dim target = If(setId = "", Nothing, Configuration.RuleSets.Single(Function(r) r.Id = setId))
                           For Each row In additions
                               Dim person = Copy(row.Person)
                               PutRecipient(person)
                               If row.Add AndAlso target IsNot Nothing AndAlso person.Active Then target.AllowedRecipientIds.Add(person.Id)
                           Next
                       End Sub)
            Return additions.Count
        End Function

        Public Function RecipientReferences(id As System.String) As System.String
            Return System.String.Join(", ", Configuration.RuleSets.Where(Function(r) r.AllowedRecipientIds.Contains(id) OrElse r.DefaultRecipientId = id OrElse r.AdditionalRules.Any(Function(a) a.RecipientIds.Contains(id))).Select(Function(r) r.Name))
        End Function

        Public Sub PutRecipient(person As RegisteredRecipient)
            If System.String.IsNullOrWhiteSpace(person.Name) Then Throw New System.IO.InvalidDataException("Enter a recipient name.")
            person = Copy(person)
            person.Name = person.Name.Trim()
            person.SmtpAddress = Registry.Address(person.SmtpAddress)
            If Configuration.Recipients.Any(Function(p) p.Id <> person.Id AndAlso p.SmtpAddress.Equals(person.SmtpAddress, System.StringComparison.OrdinalIgnoreCase)) Then Throw New System.IO.InvalidDataException("This email address is already registered.")
            Dim references = RecipientReferences(person.Id)
            If Not person.Active AndAlso references <> "" Then Throw New System.IO.InvalidDataException("Remove this recipient from these recipient sets before disabling: " & references)
            Dim index = Configuration.Recipients.FindIndex(Function(p) p.Id = person.Id)
            If index < 0 Then
                Configuration.Recipients.Add(person)
            Else
                Configuration.Recipients(index) = person
            End If
        End Sub

        Public Sub DeleteRecipient(id As System.String)
            Dim references = RecipientReferences(id)
            If references <> "" Then Throw New System.IO.InvalidDataException("Remove this recipient from these recipient sets first: " & references)
            Configuration.Recipients.RemoveAll(Function(p) p.Id = id)
        End Sub

        Public Sub PutRule(rule As RuleSet)
            Dim candidate = Copy(Configuration)
            Dim index = candidate.RuleSets.FindIndex(Function(r) r.Id = rule.Id)
            If index < 0 Then
                candidate.RuleSets.Add(Copy(rule))
            Else
                candidate.RuleSets(index) = Copy(rule)
            End If
            Registry.Validate(candidate)
            Configuration = candidate
        End Sub

        Public Shared Function ComposeName(parts As System.Collections.Generic.IEnumerable(Of System.String), separator As System.String) As System.String
            Dim pattern = System.String.Join(separator, parts)
            AttachmentNaming.ValidateTemplate(pattern)
            Return pattern
        End Function

        Public Shared Function PreviewNames(rule As RuleSet) As System.Collections.Generic.List(Of AttachmentName)
            ' A user-requested preview uses the real naming engine, including collision handling.
            Dim operation As New Operation With {.Id = "filename-preview", .ReceivedUtc = System.DateTime.UtcNow, .Extraction = New Extraction()}
            Dim suggestions As New NamingResponse()
            For index As System.Int32 = 1 To 2
                Dim id = "preview-" & index.ToString(System.Globalization.CultureInfo.InvariantCulture)
                operation.Documents.Add(New Document With {.Id = id, .FileName = "scan.pdf"})
                operation.Extraction.Sources.Add(New Source With {.Id = id, .Certain = True, .DateIso = System.DateTime.Today.ToString("yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture), .Evidence = New Evidence With {.DocumentId = id}})
                suggestions.Documents.Add(New NamingSuggestion With {.DocumentId = id, .Content = "Invoice", .DateSourceId = id})
            Next
            Return AttachmentNaming.Create(rule, operation, suggestions)
        End Function
    End Class

    Public Class AdminImportRow
        Public Property Person As RegisteredRecipient
        Public Property Line As System.Int64
        Public Property Name As System.String = ""
        Public Property Email As System.String = ""
        Public Property Organization As System.String = ""
        Public Property Status As System.String = ""
        Public Property Issue As System.String = ""
        Public Property Add As System.Boolean
        Public Property Update As System.Boolean
    End Class

    Public Class AdminSearchProgress
        Public Property Mailbox As System.String
        Public Property ThroughUtc As System.DateTime?
        Public Property Issue As System.String = ""
    End Class

    Public Class AdminActivity
        Public Property Id As System.String
        Public Property Received As System.DateTime
        Public Property Mailbox As System.String
        Public Property RuleSet As System.String
        Public Property Status As System.String
        Public Property Deliveries As System.String
        Public Property Issue As System.String
        Public Property Operation As Operation
    End Class
End Namespace
