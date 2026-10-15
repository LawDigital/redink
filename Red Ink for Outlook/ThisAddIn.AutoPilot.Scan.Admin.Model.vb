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

        Public Function PreviewRecipients(text As System.String, delimiter As System.String, hasHeader As System.Boolean) As System.Collections.Generic.List(Of AdminImportRow)
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
                    parser.TrimWhiteSpace = True
                    If hasHeader AndAlso Not parser.EndOfData Then parser.ReadFields()
                    While Not parser.EndOfData
                        If rows.Count >= 1000 Then Throw New System.IO.InvalidDataException("Import at most 1,000 rows at a time.")
                        Dim row As New AdminImportRow With {.Line = parser.LineNumber}
                        Try
                            Dim fields = parser.ReadFields()
                            If fields Is Nothing Then Exit While
                            If fields.Length < 2 OrElse fields.Length > 3 Then Throw New System.IO.InvalidDataException("Expected Name, Email and optionally Organization.")
                            row.Name = fields(0).Trim()
                            row.Email = fields(1).Trim()
                            row.Organization = If(fields.Length = 3, fields(2).Trim(), "")
                            If row.Name = "" OrElse row.Name.Any(Function(c) System.Char.IsControl(c)) OrElse row.Organization.Any(Function(c) System.Char.IsControl(c)) Then Throw New System.IO.InvalidDataException("Enter a name and keep each field on one line.")
                            row.Email = Registry.Address(row.Email)
                            If Configuration.RuleSets.Any(Function(r) r.ScannerSenders.Any(Function(a) a.Equals(row.Email, System.StringComparison.OrdinalIgnoreCase))) Then Throw New System.IO.InvalidDataException("A scanner sender cannot be a recipient.")
                            If Configuration.Recipients.Any(Function(p) p.SmtpAddress.Equals(row.Email, System.StringComparison.OrdinalIgnoreCase)) Then
                                row.Status = "Already registered — unchanged"
                            ElseIf Not seen.Add(row.Email) Then
                                row.Status = "Duplicate in import — skipped"
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

        Public Function ImportRecipients(text As System.String, delimiter As System.String, hasHeader As System.Boolean, setId As System.String) As System.Int32
            Dim rows = PreviewRecipients(text, delimiter, hasHeader)
            If rows.Any(Function(r) r.Issue <> "") Then Throw New System.IO.InvalidDataException("Correct the marked rows and preview again. Nothing has been imported.")
            Dim additions = rows.Where(Function(r) r.Add).ToList()
            If additions.Count = 0 Then Throw New System.IO.InvalidDataException("No new recipients to import.")
            SaveChange(Sub()
                           Dim target = If(setId = "", Nothing, Configuration.RuleSets.Single(Function(r) r.Id = setId))
                           For Each row In additions
                               Dim person As New RegisteredRecipient With {.Name = row.Name, .SmtpAddress = row.Email, .Organization = row.Organization}
                               PutRecipient(person)
                               If target IsNot Nothing Then target.AllowedRecipientIds.Add(person.Id)
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
        Public Property Line As System.Int64
        Public Property Name As System.String = ""
        Public Property Email As System.String = ""
        Public Property Organization As System.String = ""
        Public Property Status As System.String = ""
        Public Property Issue As System.String = ""
        Public Property Add As System.Boolean
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
