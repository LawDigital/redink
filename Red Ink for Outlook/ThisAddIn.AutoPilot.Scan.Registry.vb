' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Registry.vb
' Purpose: Versioned, locked and recoverable scanner register.
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
Imports Newtonsoft.Json

Namespace ScanInbox
    Public NotInheritable Class Registry
        Private Sub New()
        End Sub

        ' Only the model boundary accepts a complete Markdown JSON fence. Disk contracts stay strict.
        Public Shared Function ParseModel(Of T)(response As System.String) As T
            Dim json = If(response, "").Trim().TrimStart(ChrW(&HFEFF)).Trim()
            If json.StartsWith("```", System.StringComparison.Ordinal) Then
                Dim match = System.Text.RegularExpressions.Regex.Match(json, "\A```(?:json)?[ \t]*\r?\n(?<json>[\s\S]*?)\r?\n```\z", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant)
                If Not match.Success Then Throw New System.IO.InvalidDataException("Invalid fenced scan JSON response.")
                json = match.Groups("json").Value.Trim()
            End If
            Return ParseCore(Of T)(json, True)
        End Function

        Public Shared Function Parse(Of T)(json As System.String) As T
            Return ParseCore(Of T)(json, False)
        End Function

        Private Shared Function ParseCore(Of T)(json As System.String, modelOutput As System.Boolean) As T
            ' Unknown members (including model-supplied address fields) are contract errors.
            Dim settings As New Newtonsoft.Json.JsonSerializerSettings With {
                .MissingMemberHandling = Newtonsoft.Json.MissingMemberHandling.Error,
                .TypeNameHandling = Newtonsoft.Json.TypeNameHandling.None,
                .ObjectCreationHandling = Newtonsoft.Json.ObjectCreationHandling.Replace,
                .DateParseHandling = Newtonsoft.Json.DateParseHandling.None,
                .MaxDepth = 64}
            Using input As New System.IO.StringReader(json)
                Using reader As New Newtonsoft.Json.JsonTextReader(input) With {.DateParseHandling = Newtonsoft.Json.DateParseHandling.None, .MaxDepth = 64}
                    Dim token = Newtonsoft.Json.Linq.JToken.Load(reader, New Newtonsoft.Json.Linq.JsonLoadSettings With {
                        .DuplicatePropertyNameHandling = Newtonsoft.Json.Linq.DuplicatePropertyNameHandling.Error})
                    If reader.Read() Then Throw New System.IO.InvalidDataException("Unexpected trailing JSON content.")
                    Dim serializer = Newtonsoft.Json.JsonSerializer.Create(settings)
                    If modelOutput Then NormalizeModelToken(token, GetType(T), serializer)
                    Return token.ToObject(Of T)(serializer)
                End Using
            End Using
        End Function

        ' Normalize only model responses, never persisted configuration or operation files.
        ' Narrative metadata cannot grant identity, evidence, dates or transport authority.
        Private Shared Sub NormalizeTextList(owner As Newtonsoft.Json.Linq.JObject, key As System.String,
                                             diagnostics As System.Collections.Generic.List(Of System.String), Optional structured As System.Boolean = False)
            Dim prop = owner.Property(key, System.StringComparison.OrdinalIgnoreCase)
            If prop Is Nothing Then Return
            Dim original = prop.Value
            Dim values = TryCast(original, Newtonsoft.Json.Linq.JArray)
            If values Is Nothing Then
                values = New Newtonsoft.Json.Linq.JArray()
                If original.Type <> Newtonsoft.Json.Linq.JTokenType.Null Then values.Add(original.DeepClone())
            End If
            Dim normalized As New Newtonsoft.Json.Linq.JArray()
            For Each item In values
                If item.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                    normalized.Add(item.DeepClone())
                ElseIf structured Then
                    normalized.Add(item.ToString(Newtonsoft.Json.Formatting.None))
                Else
                    Dim wrapper = TryCast(item, Newtonsoft.Json.Linq.JObject)
                    Dim text = If(wrapper Is Nothing, Nothing, wrapper.Property("Text", System.StringComparison.OrdinalIgnoreCase))
                    If wrapper IsNot Nothing AndAlso wrapper.Properties().Count() = 1 AndAlso text IsNot Nothing AndAlso text.Value.Type = Newtonsoft.Json.Linq.JTokenType.String Then
                        normalized.Add(text.Value.DeepClone())
                    Else
                        Throw New System.IO.InvalidDataException("Invalid text item at " & prop.Path & "; expected text or a single Text field.")
                    End If
                End If
            Next
            If Not Newtonsoft.Json.Linq.JToken.DeepEquals(original, normalized) Then
                diagnostics.Add("Model structure normalized: " & prop.Path)
                prop.Value = normalized
            End If
        End Sub

        Private Shared Sub NormalizeModelToken(token As Newtonsoft.Json.Linq.JToken, target As System.Type, serializer As Newtonsoft.Json.JsonSerializer)
            Dim root = TryCast(token, Newtonsoft.Json.Linq.JObject)
            If root Is Nothing Then Return
            Dim diagnostics As New System.Collections.Generic.List(Of System.String)()
            If target Is GetType(Extraction) Then
                NormalizeTextList(root, "Problems", diagnostics)
                Dim sources = TryCast(root.GetValue("Sources", System.StringComparison.OrdinalIgnoreCase), Newtonsoft.Json.Linq.JArray)
                If sources IsNot Nothing Then
                    For Each source In sources.OfType(Of Newtonsoft.Json.Linq.JObject)()
                        NormalizeTextList(source, "Relationships", diagnostics, structured:=True)
                        For Each key In New System.String() {"Subject", "SubjectEvidence", "DateIso", "DateEvidence", "DocumentType", "Language"}
                            Dim prop = source.Property(key, System.StringComparison.OrdinalIgnoreCase)
                            If prop Is Nothing Then Continue For
                            Try
                                If key.EndsWith("Evidence", System.StringComparison.Ordinal) Then
                                    prop.Value.ToObject(Of Evidence)(serializer)
                                ElseIf prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then
                                    prop.Value = New Newtonsoft.Json.Linq.JValue("")
                                ElseIf prop.Value.Type <> Newtonsoft.Json.Linq.JTokenType.String Then
                                    Throw New System.IO.InvalidDataException("Expected optional text.")
                                End If
                            Catch ex As System.Exception When TypeOf ex Is Newtonsoft.Json.JsonException OrElse TypeOf ex Is System.IO.InvalidDataException
                                diagnostics.Add("Optional source metadata omitted: " & prop.Path & "; " & ex.Message & " Original: " & prop.Value.ToString(Newtonsoft.Json.Formatting.None))
                                prop.Value = If(key.EndsWith("Evidence", System.StringComparison.Ordinal), Newtonsoft.Json.Linq.JValue.CreateNull(), New Newtonsoft.Json.Linq.JValue(""))
                            End Try
                        Next
                        For Each key In New System.String() {"Summary", "Actions", "Uncertainties"}
                            Try
                                NormalizeTextList(source, key, diagnostics)
                            Catch ex As System.IO.InvalidDataException
                                Dim prop = source.Property(key, System.StringComparison.OrdinalIgnoreCase)
                                diagnostics.Add("Optional narrative omitted: " & ex.Message & " Original: " & prop.Value.ToString(Newtonsoft.Json.Formatting.None))
                                prop.Value = New Newtonsoft.Json.Linq.JArray()
                            End Try
                        Next
                    Next
                End If
                ' A malformed optional calendar claim must not discard validated address blocks.
                For Each key In New System.String() {"Deadlines", "Stamps", "RuleMatches"}
                    Dim prop = root.Property(key, System.StringComparison.OrdinalIgnoreCase)
                    If prop Is Nothing Then Continue For
                    If prop.Value.Type = Newtonsoft.Json.Linq.JTokenType.Null Then
                        prop.Value = New Newtonsoft.Json.Linq.JArray()
                        diagnostics.Add("Empty optional collection normalized: " & key)
                    End If
                    Dim entries = TryCast(prop.Value, Newtonsoft.Json.Linq.JArray)
                    If entries Is Nothing Then Continue For
                    For Each entry In entries.ToList()
                        Try
                            Dim kind = If(key = "Deadlines", GetType(Deadline), If(key = "Stamps", GetType(ReceiptStamp), GetType(RuleMatch)))
                            If entry.Type <> Newtonsoft.Json.Linq.JTokenType.Object Then Throw New System.IO.InvalidDataException("Expected an object.")
                            entry.ToObject(kind, serializer)
                        Catch ex As System.Exception When TypeOf ex Is Newtonsoft.Json.JsonException OrElse TypeOf ex Is System.IO.InvalidDataException
                            diagnostics.Add("Optional " & key & " entry omitted: " & ex.Message & " Original: " & entry.ToString(Newtonsoft.Json.Formatting.None))
                            entry.Remove()
                        End Try
                    Next
                Next
                Dim problems = TryCast(root.GetValue("Problems", System.StringComparison.OrdinalIgnoreCase), Newtonsoft.Json.Linq.JArray)
                If problems IsNot Nothing Then
                    For Each diagnostic In diagnostics
                        problems.Add(diagnostic)
                    Next
                End If
            ElseIf target Is GetType(PersonalTaskReport) Then
                Dim tasks = TryCast(root.GetValue("Tasks", System.StringComparison.OrdinalIgnoreCase), Newtonsoft.Json.Linq.JArray)
                If tasks IsNot Nothing Then
                    For Each task In tasks.OfType(Of Newtonsoft.Json.Linq.JObject)()
                        NormalizeTextList(task, "Paragraphs", diagnostics)
                    Next
                End If
            ElseIf target Is GetType(Presentation) Then
                For Each key In New System.String() {"Summary", "Actions", "EventTitles", "ReviewNotes"}
                    NormalizeTextList(root, key, diagnostics)
                Next
            End If
        End Sub

        Public Shared Function Load(path As System.String) As Configuration
            If Not System.IO.File.Exists(path) Then Throw New System.IO.FileNotFoundException("Scan register is missing; delivery stopped.", path)
            Dim value = Parse(Of Configuration)(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8))
            Validate(value)
            Return value
        End Function

        Public Shared Function Hash(text As System.String) As System.String
            Using sha = System.Security.Cryptography.SHA256.Create()
                Return System.BitConverter.ToString(sha.ComputeHash(System.Text.Encoding.UTF8.GetBytes(text))).Replace("-", "").ToLowerInvariant()
            End Using
        End Function

        Public Shared Function DeliveryAuthorityHash(config As Configuration) As System.String
            Dim snapshot = Parse(Of Configuration)(Newtonsoft.Json.JsonConvert.SerializeObject(config))
            snapshot.Revision = 0
            For Each person In snapshot.Recipients
                person.Instruction = ""
                person.MailLanguage = ""
                person.TranslateAttachments = False
                person.TranslationLanguage = ""
                person.DraftReply = False
                person.WebResearch = False
                person.TaskInstructions = ""
            Next
            ' Keep the prior authority shape: this personal option grants no routing permission.
            Dim authority = Newtonsoft.Json.Linq.JObject.FromObject(snapshot)
            For Each person As Newtonsoft.Json.Linq.JObject In DirectCast(authority("Recipients"), Newtonsoft.Json.Linq.JArray)
                person.Remove("ReplyOnlyWhenAppropriate")
            Next
            Return Hash(authority.ToString(Newtonsoft.Json.Formatting.None))
        End Function

        Public Shared Function MatchesPrePersonalConfiguration(config As Configuration, operation As Operation) As System.Boolean
            ' Older pending operations hashed the same register without the new preference fields.
            ' Verify every pre-existing field exactly, at the captured revision, before adopting authority.
            If operation.AuthorityHash <> "" OrElse operation.ConfigurationHash = "" Then Return False
            Dim snapshot = Newtonsoft.Json.Linq.JObject.FromObject(config)
            snapshot("Revision") = New Newtonsoft.Json.Linq.JValue(operation.ConfigurationRevision)
            For Each person As Newtonsoft.Json.Linq.JObject In DirectCast(snapshot("Recipients"), Newtonsoft.Json.Linq.JArray)
                For Each key In New System.String() {"MailLanguage", "TranslateAttachments", "TranslationLanguage", "DraftReply", "ReplyOnlyWhenAppropriate", "WebResearch", "TaskInstructions"}
                    person.Remove(key)
                Next
            Next
            Return Hash(snapshot.ToString(Newtonsoft.Json.Formatting.None)) = operation.ConfigurationHash
        End Function

        Public Shared Function Address(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then Throw New System.IO.InvalidDataException("An SMTP address is empty.")
            Dim parsed As New System.Net.Mail.MailAddress(value.Trim())
            If Not parsed.Address.Equals(value.Trim(), System.StringComparison.OrdinalIgnoreCase) OrElse Not parsed.Address.Contains("@") Then
                Throw New System.IO.InvalidDataException("Use one bare SMTP address per entry.")
            End If
            Return parsed.Address.ToLowerInvariant()
        End Function

        Private Shared Sub Require(condition As System.Boolean, message As System.String)
            If Not condition Then Throw New System.IO.InvalidDataException(message)
        End Sub

        Private Shared Sub UniqueIds(ids As System.Collections.Generic.IEnumerable(Of System.String), label As System.String)
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            For Each id In ids
                Require(Not System.String.IsNullOrWhiteSpace(id) AndAlso id.Length <= 100 AndAlso seen.Add(id), "Missing/duplicate " & label & " ID.")
            Next
        End Sub

        Public Shared Sub Validate(config As Configuration)
            Require(config IsNot Nothing AndAlso config.SchemaVersion = 1, "Unsupported scan register version.")
            Require(config.Revision >= 0 AndAlso config.InitialLookbackDays >= 1 AndAlso config.InitialLookbackDays <= 365 AndAlso config.OverlapHours >= 1 AndAlso config.OverlapHours <= 168, "Invalid scan lookback/revision.")
            Require(config.Recipients IsNot Nothing AndAlso config.RuleSets IsNot Nothing, "Missing register collections.")
            Require(config.Recipients.All(Function(p) p IsNot Nothing) AndAlso config.RuleSets.All(Function(p) p IsNot Nothing), "Null register entries.")
            UniqueIds(config.Recipients.Select(Function(p) p.Id), "recipient")
            UniqueIds(config.RuleSets.Select(Function(p) p.Id), "rule set")
            Dim smtp As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
            For Each person In config.Recipients
                Require(Not System.String.IsNullOrWhiteSpace(person.Name), "Recipient name is required.")
                Require(person.Aliases IsNot Nothing AndAlso person.Attributes IsNot Nothing, "Recipient aliases/attributes must be arrays.")
                ValidatePersonalSettings(person)
                person.SmtpAddress = Address(person.SmtpAddress)
                ValidateCopyRecipients(person)
                Require(smtp.Add(person.SmtpAddress), "Duplicate register SMTP address.")
            Next
            Dim combinations As New System.Collections.Generic.List(Of System.Tuple(Of System.String, System.String))()
            For Each rule In config.RuleSets
                Require(Not System.String.IsNullOrWhiteSpace(rule.Name), "Rule set name is required.")
                Require(rule.ScannerSenders IsNot Nothing AndAlso rule.ScannerSenders.Count > 0 AndAlso rule.AllowedRecipientIds IsNot Nothing AndAlso rule.AdditionalRules IsNot Nothing AndAlso rule.ReminderDays IsNot Nothing, "Incomplete rule set.")
                System.TimeZoneInfo.FindSystemTimeZoneById(rule.TimeZoneId)
                AttachmentNaming.ValidateTemplate(rule.AttachmentNameTemplate)
                ValidateSubjectTemplate(rule.SubjectTemplate)
                Require(rule.AttachmentNamingInstruction Is Nothing OrElse rule.AttachmentNamingInstruction.Length <= 20000, "Attachment naming instruction is too long.")
                Require(rule.DeadlineProfile = "calendar_days" OrElse rule.DeadlineProfile = "review_only", "Unsupported deadline profile; choose calendar_days or review_only.")
                Require(rule.RetentionDays >= 1 AndAlso rule.RetentionDays <= 3650, "RetentionDays must be 1..3650.")
                Require(rule.ReminderDays.Count > 0 AndAlso rule.ReminderDays.All(Function(d) d > 0 AndAlso d <= 3650) AndAlso rule.ReminderDays.Distinct().Count() = rule.ReminderDays.Count, "Invalid reminder days.")
                UniqueIds(rule.AllowedRecipientIds, "allowed recipient")
                For Each id In rule.AllowedRecipientIds
                    Require(config.Recipients.Any(Function(p) p.Id = id AndAlso p.Active), "Allowed recipient missing or inactive: " & id)
                Next
                Require(rule.AllowedRecipientIds.Contains(rule.DefaultRecipientId), "Default must be active and allowed.")
                UniqueIds(rule.AdditionalRules.Select(Function(r) r.Id), "additional rule")
                For Each extra In rule.AdditionalRules
                    Require(extra IsNot Nothing AndAlso (extra.Kind = "exact" OrElse extra.Kind = "semantic") AndAlso Not System.String.IsNullOrWhiteSpace(extra.Criteria), "Invalid additional rule.")
                    Require(extra.RecipientIds IsNot Nothing AndAlso extra.RecipientIds.Count > 0 AndAlso extra.RecipientIds.All(Function(id) rule.AllowedRecipientIds.Contains(id)), "Additional recipients must be allowed.")
                Next
                Dim mailbox = If(System.String.IsNullOrWhiteSpace(rule.Mailbox), "", Address(rule.Mailbox))
                rule.Mailbox = mailbox
                rule.ScannerSenders = rule.ScannerSenders.Select(Function(sender) Address(sender)).ToList()
                Dim ownSenders As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                For Each sender In rule.ScannerSenders
                    Dim normalized = Address(sender)
                    Require(ownSenders.Add(normalized), "Duplicate scanner sender in rule set.")
                    Require(Not config.Recipients.Any(Function(p) Address(p.SmtpAddress) = normalized OrElse p.CopySmtpAddresses.Contains(normalized)), "Scanner sender cannot also be a delivery or copy recipient.")
                    If rule.Active Then
                        Require(Not combinations.Any(Function(c) c.Item1 = normalized AndAlso (c.Item2 = mailbox OrElse c.Item2 = "" OrElse mailbox = "")), "Ambiguous scanner sender/mailbox combination.")
                        combinations.Add(System.Tuple.Create(normalized, mailbox))
                    End If
                Next
            Next
        End Sub

        Public Shared Sub ValidateCopyRecipients(person As RegisteredRecipient)
            Require(person.CopySmtpAddresses IsNot Nothing, "Copy email addresses must be an array.")
            person.CopySmtpAddresses = person.CopySmtpAddresses.Select(Function(value) Address(value)).
                Distinct(System.StringComparer.OrdinalIgnoreCase).ToList()
        End Sub

        ' Expand only the participants in this actual delivery (including personalised and split
        ' deliveries). Copies never recursively expand other recipients or run their personal tasks.
        ' True means To, False means CC; To always wins across duplicate SMTP addresses.
        Public Shared Function DeliveryAddresses(config As Configuration, rule As RuleSet, delivery As Delivery,
                                                 mailbox As System.String) As System.Collections.Generic.Dictionary(Of System.String, System.Boolean)
            Require(delivery.ToIds IsNot Nothing AndAlso delivery.ToIds.Count > 0 AndAlso delivery.CcIds IsNot Nothing, "Scan delivery requires recipients.")
            Dim result As New System.Collections.Generic.Dictionary(Of System.String, System.Boolean)(System.StringComparer.OrdinalIgnoreCase)
            Dim people As New System.Collections.Generic.List(Of RegisteredRecipient)()
            For Each id In delivery.ToIds.Concat(delivery.CcIds).Distinct()
                Dim person = config.Recipients.SingleOrDefault(Function(p) p.Active AndAlso p.Id = id AndAlso rule.AllowedRecipientIds.Contains(id))
                Require(person IsNot Nothing, "Scan delivery recipient is missing, inactive or not allowed: " & id)
                people.Add(person)
                Dim smtp = Address(person.SmtpAddress)
                result(smtp) = delivery.ToIds.Contains(id) OrElse (result.ContainsKey(smtp) AndAlso result(smtp))
            Next
            For Each person In people
                For Each value In person.CopySmtpAddresses
                    Dim smtp = Address(value)
                    Require(Not smtp.Equals(mailbox, System.StringComparison.OrdinalIgnoreCase) AndAlso
                        Not config.RuleSets.Any(Function(r) r.ScannerSenders.Contains(smtp, System.StringComparer.OrdinalIgnoreCase)),
                        "A scan copy recipient cannot be the receiving mailbox or a scanner sender.")
                    If Not result.ContainsKey(smtp) Then result.Add(smtp, False)
                Next
            Next
            Return result
        End Function

        Public Shared Sub ValidatePersonalSettings(person As RegisteredRecipient)
            For Each language In New System.String() {person.MailLanguage, person.TranslationLanguage}
                Require(language IsNot Nothing AndAlso language.Length <= 80 AndAlso Not language.Any(Function(c) System.Char.IsControl(c)), "Language must be a single line of at most 80 characters.")
            Next
            Require(Not person.TranslateAttachments OrElse Not System.String.IsNullOrWhiteSpace(person.TranslationLanguage), "Choose a target language for attachment translations.")
            Require(person.Instruction IsNot Nothing AndAlso person.Instruction.Length <= 20000 AndAlso person.TaskInstructions IsNot Nothing AndAlso person.TaskInstructions.Length <= 20000, "Personal instructions must contain at most 20,000 characters.")
        End Sub

        Public Shared Function PersonalSettings(path As System.String, trustedSender As System.String, action As System.String, changes As Newtonsoft.Json.Linq.JObject) As Newtonsoft.Json.Linq.JObject
            ' The caller supplies the transport sender, never a model-selected identity.
            Dim sender = Address(trustedSender)
            Dim bytes = System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)
            Dim config = Parse(Of Configuration)(bytes)
            Validate(config)
            Dim person = config.Recipients.SingleOrDefault(Function(p) p.Active AndAlso p.SmtpAddress.Equals(sender, System.StringComparison.OrdinalIgnoreCase))
            Require(person IsNot Nothing, "This sender is not an active registered scan recipient.")
            Dim keys = New System.String() {"Instruction", "MailLanguage", "TranslateAttachments", "TranslationLanguage", "DraftReply", "ReplyOnlyWhenAppropriate", "WebResearch", "TaskInstructions"}
            Dim result As New Newtonsoft.Json.Linq.JObject()
            Dim source = Newtonsoft.Json.Linq.JObject.FromObject(person)
            For Each key In keys
                result.Add(key, source(key).DeepClone())
            Next
            Require(action = "get" OrElse action = "update", "Use get or update.")
            If action = "get" Then
                Require(changes Is Nothing OrElse Not changes.Properties().Any(), "get does not accept changes.")
                Return result
            End If
            Require(changes IsNot Nothing AndAlso changes.Properties().Any(), "Provide at least one setting to update.")
            For Each change In changes.Properties()
                Require(keys.Contains(change.Name), "Only personal scan settings can be updated; recipient identity, address and routing are administrator-controlled.")
                Require(change.Value.Type = result(change.Name).Type, "Invalid setting type: " & change.Name)
                result(change.Name) = change.Value.DeepClone()
                source(change.Name) = change.Value.DeepClone()
            Next
            Dim updated = source.ToObject(Of RegisteredRecipient)()
            ValidatePersonalSettings(updated)
            config.Recipients(config.Recipients.IndexOf(person)) = updated
            Save(path, config, config.Revision, Hash(bytes))
            Return result
        End Function

        Public Shared Sub ValidateSubjectTemplate(value As System.String)
            Dim template = If(value, "")
            Require(template.Length <= 240 AndAlso Not template.Any(Function(c) System.Char.IsControl(c)), "Email subject must be one line of at most 240 characters.")
            Dim literal = template.Replace("{subject}", "")
            Require(Not literal.Contains("{") AndAlso Not literal.Contains("}"), "Only {subject} is supported in the email subject.")
        End Sub

        Public Shared Function EmailSubject(template As System.String, description As System.String) As System.String
            ValidateSubjectTemplate(template)
            Dim subject = If(System.String.IsNullOrWhiteSpace(template), "Scan: {subject}", template.Trim()).Replace("{subject}", If(description, "").Trim())
            Require(Not System.String.IsNullOrWhiteSpace(subject) AndAlso subject.Length <= 512 AndAlso Not subject.Any(Function(c) System.Char.IsControl(c)), "Invalid generated email subject.")
            Return subject
        End Function

        Public Shared Function Match(config As Configuration, sender As System.String, mailbox As System.String) As RuleSet
            Dim normalized = Address(sender)
            Dim found = config.RuleSets.Where(Function(r) r.Active AndAlso r.ScannerSenders.Any(Function(s) Address(s) = normalized) AndAlso
                (System.String.IsNullOrWhiteSpace(r.Mailbox) OrElse Address(r.Mailbox) = mailbox.ToLowerInvariant())).ToList()
            If found.Count > 1 Then Throw New System.IO.InvalidDataException("Ambiguous scanner classification.")
            Return found.SingleOrDefault()
        End Function

        Public Shared Sub AtomicWrite(path As System.String, json As System.String)
            Dim temp = path & "." & System.Guid.NewGuid().ToString("N") & ".tmp"
            Try
                Using stream As New System.IO.FileStream(temp, System.IO.FileMode.CreateNew, System.IO.FileAccess.Write, System.IO.FileShare.None)
                    Dim bytes = New System.Text.UTF8Encoding(False).GetBytes(json)
                    stream.Write(bytes, 0, bytes.Length)
                    stream.Flush(True)
                End Using
                If System.IO.File.Exists(path) Then
                    System.IO.File.Replace(temp, path, path & ".bak")
                Else
                    System.IO.File.Move(temp, path)
                End If
            Finally
                If System.IO.File.Exists(temp) Then System.IO.File.Delete(temp)
            End Try
        End Sub

        Public Shared Sub Save(path As System.String, config As Configuration, expectedRevision As System.Int64, Optional expectedContentHash As System.String = "", Optional expectedMissing As System.Boolean = False)
            Validate(config)
            Using gate As New System.IO.FileStream(path & ".lock", System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None)
                Dim exists = System.IO.File.Exists(path)
                If expectedMissing Then Require(Not exists, "The scan register was created in another window. Reload before saving.")
                Dim actualHash = If(exists, Hash(System.IO.File.ReadAllText(path, System.Text.Encoding.UTF8)), "")
                If expectedContentHash <> "" Then Require(actualHash = expectedContentHash, "The scan register changed. Reload before saving.")
                Dim currentRevision As System.Int64 = 0
                If exists Then
                    Try
                        currentRevision = Load(path).Revision
                    Catch ex As System.Exception
                        ' Explicit reviewed repair is permitted only against the exact loaded bytes.
                        If expectedContentHash = "" OrElse actualHash <> expectedContentHash Then Throw
                        currentRevision = expectedRevision
                        config.Revision = System.Math.Max(config.Revision, expectedRevision)
                    End Try
                End If
                Require(currentRevision = expectedRevision, "The scan register changed. Reload before saving.")
                config.Revision = System.Math.Max(currentRevision, config.Revision) + 1
                AtomicWrite(path, Newtonsoft.Json.JsonConvert.SerializeObject(config, Newtonsoft.Json.Formatting.Indented))
            End Using
        End Sub
    End Class
End Namespace
