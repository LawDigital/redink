' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Decisions.vb
' Purpose: Validate model recipient choices; deterministic deadlines and calendar serialization.
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
    Public NotInheritable Class Decisions
        Public Const RecipientMatchThreshold As System.Double = 0.8R

        Private Sub New()
        End Sub

        Private Shared Sub Check(value As System.Boolean, message As System.String)
            If Not value Then Throw New System.IO.InvalidDataException("Scan extraction contract: " & message)
        End Sub

        Public Shared Function IsoDate(value As System.String) As System.DateTime?
            Dim parsed As System.DateTime
            If System.DateTime.TryParseExact(value, "yyyy-MM-dd", System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, parsed) Then Return parsed.Date
            Return Nothing
        End Function

        Public Shared Function Normalize(value As System.String) As System.String
            Return System.Text.RegularExpressions.Regex.Replace(If(value, ""), "\s+", " ").Trim().ToUpperInvariant()
        End Function

        Public Shared Function ContainsLiteral(text As System.String, literal As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(literal) Then Return False
            Return System.Text.RegularExpressions.Regex.IsMatch(Normalize(text), "(?<![\p{L}\p{N}])" &
                System.Text.RegularExpressions.Regex.Escape(Normalize(literal)) & "(?![\p{L}\p{N}])", System.Text.RegularExpressions.RegexOptions.CultureInvariant)
        End Function

        Public Shared Sub ValidateEvidence(evidence As Evidence, documents As System.Collections.Generic.List(Of Document))
            Check(evidence IsNot Nothing, "Missing evidence.")
            Dim doc = documents.SingleOrDefault(Function(d) d.Id = evidence.DocumentId)
            Check(doc IsNot Nothing, "Unknown evidence document.")
            Check(evidence.StartPage >= 1 AndAlso evidence.EndPage >= evidence.StartPage AndAlso evidence.EndPage <= doc.Pages.Count, "Invalid evidence page range.")
            Check(Not System.String.IsNullOrWhiteSpace(evidence.Quote), "Empty evidence quote.")
            Dim text = System.String.Join(System.Environment.NewLine, doc.Pages.Skip(evidence.StartPage - 1).Take(evidence.EndPage - evidence.StartPage + 1))
            Check(Normalize(text).Contains(Normalize(evidence.Quote)), "Evidence quote not found in the declared pages.")
        End Sub

        Public Shared Sub ValidateExtraction(value As Extraction, documents As System.Collections.Generic.List(Of Document), config As Configuration, rule As RuleSet)
            Check(value IsNot Nothing AndAlso value.Sources IsNot Nothing AndAlso value.Deadlines IsNot Nothing AndAlso value.Stamps IsNot Nothing AndAlso value.RuleMatches IsNot Nothing AndAlso value.Problems IsNot Nothing, "Missing extraction collections.")
            Check(value.Sources.All(Function(s) s IsNot Nothing) AndAlso value.Sources.Select(Function(s) s.Id).Distinct().Count() = value.Sources.Count, "Duplicate/null sources.")
            For Each source In value.Sources
                Check(Not System.String.IsNullOrWhiteSpace(source.Id), "Missing source ID.")
                ValidateEvidence(source.Evidence, documents)
                ValidateEvidence(source.ContextEvidence, documents)
                source.FunctionInPackage = If(source.FunctionInPackage, "").Trim().ToLowerInvariant()
                If Not {"current_cover", "current_letter", "attachment", "historical", "unknown"}.Contains(source.FunctionInPackage) Then
                    value.Problems.Add("Unknown source function: " & source.FunctionInPackage & "; default routing required.")
                    source.FunctionInPackage = "unknown"
                    source.Certain = False
                End If
                Check(source.Recipients IsNot Nothing AndAlso source.Uncertainties IsNot Nothing AndAlso source.Relationships IsNot Nothing AndAlso source.Summary IsNot Nothing AndAlso source.Actions IsNot Nothing, "Missing source arrays.")
                If source.DateIso <> "" Then
                    Try
                        Check(IsoDate(source.DateIso).HasValue, "Invalid document date.")
                        ValidateEvidence(source.DateEvidence, documents)
                        Check(InSource(source.DateEvidence, source), "Document date outside source.")
                    Catch ex As System.IO.InvalidDataException
                        value.Problems.Add("Optional letter date discarded: " & ex.Message)
                        source.DateIso = ""
                        source.DateEvidence = Nothing
                    End Try
                End If
                If Not System.String.IsNullOrWhiteSpace(source.Subject) Then
                    Try
                        Check(source.Subject.Length <= 2000 AndAlso Not source.Subject.Any(Function(c) System.Char.IsControl(c)), "Invalid document subject.")
                        ValidateEvidence(source.SubjectEvidence, documents)
                        Check(InSource(source.SubjectEvidence, source) AndAlso Normalize(source.SubjectEvidence.Quote).Contains(Normalize(source.Subject)), "Document subject is not supported by source evidence.")
                    Catch ex As System.IO.InvalidDataException
                        ' Optional presentation metadata cannot invalidate recipient/deadline evidence.
                        value.Problems.Add("Document heading discarded: " & ex.Message)
                        source.Subject = ""
                        source.SubjectEvidence = Nothing
                    End Try
                End If
                For Each recipient In source.Recipients
                    Check(recipient IsNot Nothing AndAlso recipient.CandidateIds IsNot Nothing, "Null recipient/IDs.")
                    recipient.Role = If(recipient.Role, "").Trim().ToLowerInvariant()
                    Select Case recipient.Role
                        Case "to", "addressee", "recipient", "primary" : recipient.Role = "direct"
                        Case "cc", "carbon_copy" : recipient.Role = "copy"
                        Case "direct", "copy", "mention", "signature"
                        Case Else
                            recipient.Role = "unknown"
                            source.Certain = False
                            source.Uncertainties.Add("An address-block role could not be established; use the default recipient.")
                    End Select
                    ValidateEvidence(recipient.Evidence, documents)
                    Check(InSource(recipient.Evidence, source), "Recipient evidence outside source.")
                    If recipient.MatchConfidence.HasValue Then
                        Check(Not System.Double.IsNaN(recipient.MatchConfidence.Value) AndAlso Not System.Double.IsInfinity(recipient.MatchConfidence.Value) AndAlso recipient.MatchConfidence.Value >= 0 AndAlso recipient.MatchConfidence.Value <= 1, "Recipient match confidence must be between 0 and 1.")
                    End If
                    If Not System.String.IsNullOrEmpty(recipient.MatchedRecipientId) Then
                        Check(rule.AllowedRecipientIds.Contains(recipient.MatchedRecipientId) AndAlso config.Recipients.Any(Function(p) p.Active AndAlso p.Id = recipient.MatchedRecipientId), "Unknown/disallowed matched recipient ID.")
                    End If
                    For Each id In recipient.CandidateIds
                        Check(rule.AllowedRecipientIds.Contains(id) AndAlso config.Recipients.Any(Function(p) p.Id = id AndAlso p.Active), "Unknown/disallowed recipient ID.")
                    Next
                Next
            Next
            Dim deadlineIds As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal)
            Dim duplicateIds = value.Deadlines.Where(Function(d) d IsNot Nothing).GroupBy(Function(d) d.Id).Where(Function(g) g.Count() > 1).Select(Function(g) g.Key).ToList()
            For Each deadline In value.Deadlines.ToList()
                Try
                    Check(deadline IsNot Nothing, "Null deadline.")
                    Check(Not System.String.IsNullOrWhiteSpace(deadline.Id) AndAlso Not duplicateIds.Contains(deadline.Id) AndAlso deadlineIds.Add(deadline.Id), "Missing/duplicate deadline ID.")
                    ValidateEvidence(deadline.Evidence, documents)
                    deadline.Status = If(deadline.Status, "").Trim().ToLowerInvariant()
                    deadline.Trigger = If(deadline.Trigger, "").Trim().ToLowerInvariant()
                    deadline.Unit = If(deadline.Unit, "").Trim().ToLowerInvariant()
                    deadline.Profile = If(deadline.Profile, "").Trim().ToLowerInvariant()
                    Check({"current", "historical", "unclear"}.Contains(deadline.Status), "Unknown deadline status.")
                    Check({"receipt", "explicit", "other", ""}.Contains(deadline.Trigger), "Unknown deadline trigger.")
                    Check(deadline.Duration >= 0 AndAlso deadline.Duration <= 36500, "Invalid deadline duration.")
                    Check(Not System.String.IsNullOrWhiteSpace(deadline.Action), "Missing deadline action.")
                    If Not System.String.IsNullOrEmpty(deadline.DateIso) Then Check(IsoDate(deadline.DateIso).HasValue, "Invalid deadline date.")
                Catch ex As System.IO.InvalidDataException
                    value.Problems.Add("Calendar claim omitted: " & ex.Message & " Original: " & Newtonsoft.Json.JsonConvert.SerializeObject(deadline))
                    value.Deadlines.Remove(deadline)
                End Try
            Next
            For Each stamp In value.Stamps.ToList()
                Try
                    Check(stamp IsNot Nothing AndAlso IsoDate(stamp.DateIso).HasValue, "Invalid receipt stamp.")
                    stamp.Relation = If(stamp.Relation, "").Trim().ToLowerInvariant()
                    Check({"current", "historical", "unclear"}.Contains(stamp.Relation), "Invalid stamp relation.")
                    ValidateEvidence(stamp.Evidence, documents)
                Catch ex As System.IO.InvalidDataException
                    value.Problems.Add("Receipt stamp omitted; email receipt basis remains available: " & ex.Message & " Original: " & Newtonsoft.Json.JsonConvert.SerializeObject(stamp))
                    value.Stamps.Remove(stamp)
                End Try
            Next
            For Each hit In value.RuleMatches.ToList()
                Try
                    Check(hit IsNot Nothing AndAlso rule.AdditionalRules.Any(Function(r) r.Active AndAlso r.Kind = "semantic" AndAlso r.Id = hit.RuleId), "Unknown semantic rule ID.")
                    ValidateEvidence(hit.Evidence, documents)
                Catch ex As System.IO.InvalidDataException
                    value.Problems.Add("Optional topic claim omitted: " & ex.Message & " Original: " & Newtonsoft.Json.JsonConvert.SerializeObject(hit))
                    value.RuleMatches.Remove(hit)
                End Try
            Next
        End Sub

        Private Shared Function InSource(evidence As Evidence, source As Source) As System.Boolean
            Return evidence.DocumentId = source.Evidence.DocumentId AndAlso evidence.StartPage >= source.Evidence.StartPage AndAlso evidence.EndPage <= source.Evidence.EndPage
        End Function

        Private Shared Function Rank(source As Source) As System.Int32
            If Not source.Certain Then Return 0
            Select Case source.FunctionInPackage
                Case "current_cover" : Return 2
                Case "current_letter" : Return 1
                Case Else : Return 0
            End Select
        End Function

        Private Shared Function SelectSource(value As Extraction, result As Decision) As Source
            Dim candidates = value.Sources.Where(Function(s) Rank(s) > 0).ToList()
            Dim selected As Source = Nothing
            Dim uncertainCurrent = value.Sources.Where(Function(s) (s.FunctionInPackage = "current_cover" OrElse s.FunctionInPackage = "current_letter" OrElse s.FunctionInPackage = "unknown") AndAlso Rank(s) = 0).ToList()
            If candidates.Count > 0 Then
                Dim best = candidates.Max(Function(s) Rank(s))
                candidates = candidates.Where(Function(s) Rank(s) = best).ToList()
                If uncertainCurrent.Any(Function(s) s.FunctionInPackage = "unknown" OrElse s.FunctionInPackage = "current_cover" OrElse best = 1) Then
                    result.Reasons.Add("Unresolved competing current source; administrative default required.")
                ElseIf candidates.Count = 1 Then
                    selected = candidates(0)
                ElseIf candidates.All(Function(s) IsoDate(s.DateIso).HasValue) Then
                    Dim latest = candidates.Max(Function(s) IsoDate(s.DateIso).Value)
                    Dim winners = candidates.Where(Function(s) IsoDate(s.DateIso).Value = latest).ToList()
                    If winners.Count = 1 Then selected = winners(0)
                End If
            End If
            If selected Is Nothing Then
                result.Reasons.Add("Default: no unique reliable current delivery source; equal/unknown dates or unresolved context.")
            Else
                result.SourceId = selected.Id
                result.Reasons.Add("Selected " & selected.Id & " (" & selected.FunctionInPackage & ", document date " & If(selected.DateIso = "", "unknown", selected.DateIso) & "). Context precedes date; dates only break equal-context ties.")
            End If
            For Each source In value.Sources
                If selected Is Nothing OrElse source.Id <> selected.Id Then
                    result.Reasons.Add("Rejected source " & source.Id & ": " & source.FunctionInPackage & "; date=" & source.DateIso & "; certainty=" & source.Certain.ToString() & "; context rank=" & Rank(source).ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If
            Next
            Return selected
        End Function

        Private Shared Function MatchPerson(claim As RecipientClaim) As System.String
            ' Identity selection belongs exclusively to the model. No lexical/name heuristic.
            ' ValidateExtraction has already checked the finite register and confidence range.
            If Not claim.MatchConfidence.HasValue OrElse claim.MatchConfidence.Value < RecipientMatchThreshold Then Return ""
            Return If(claim.MatchedRecipientId, "")
        End Function

        Public Shared Function Decide(config As Configuration, rule As RuleSet, documents As System.Collections.Generic.List(Of Document), extraction As Extraction,
                                      receivedUtc As System.DateTime, nowUtc As System.DateTime) As Decision
            ValidateExtraction(extraction, documents, config, rule)
            Dim result As New Decision()
            ' Check page coverage independently of the model's certainty claims.
            For Each document In documents
                For page As System.Int32 = 1 To document.Pages.Count
                    Dim pageNumber = page
                    If Not System.String.IsNullOrWhiteSpace(document.Pages(page - 1)) AndAlso Not extraction.Sources.Any(Function(s) s.Evidence.DocumentId = document.Id AndAlso s.Evidence.StartPage <= pageNumber AndAlso s.Evidence.EndPage >= pageNumber) Then
                        document.Complete = False
                        document.Problems.Add("No classified source for readable page/section " & page.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; analysis cannot claim completeness.")
                    End If
                Next
            Next
            Dim selected = SelectSource(extraction, result)
            Dim unresolved As System.Boolean = selected Is Nothing OrElse documents.Any(Function(d) Not d.Complete)
            If selected IsNot Nothing Then
                For Each claim In selected.Recipients.Where(Function(c) c.Role = "direct" OrElse c.Role = "copy")
                    ' External recipients retain their document role, but can never become addresses.
                    If claim.CandidateIds.Count = 0 AndAlso System.String.IsNullOrEmpty(claim.MatchedRecipientId) Then Continue For
                    Dim id = MatchPerson(claim)
                    result.Reasons.Add("Model recipient match for " & claim.Name & ": " & If(id = "", "below threshold or unresolved", config.Recipients.Single(Function(p) p.Id = id).Name) & "; confidence=" & If(claim.MatchConfidence.HasValue, claim.MatchConfidence.Value.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture), "missing") & "; " & claim.MatchReason)
                    If id = "" Then
                        unresolved = True
                    ElseIf claim.Role = "direct" Then
                        result.ToIds.Add(id)
                    Else
                        result.CopyIds.Add(id)
                    End If
                Next
            End If
            ' Routing tiers are exclusive: named addressees, topic rules, then default.
            If unresolved Then
                result.ToIds.Clear()
                result.CopyIds.Clear()
            End If
            result.DirectIds.AddRange(result.ToIds.Distinct())
            result.CcIds.AddRange(result.CopyIds)
            If result.ToIds.Count = 0 AndAlso result.CopyIds.Count > 0 Then
                result.ToIds.AddRange(result.CopyIds)
                result.CcIds.Clear()
            End If
            If Not unresolved AndAlso result.ToIds.Count = 0 Then
                Dim allText = System.String.Join(System.Environment.NewLine, documents.SelectMany(Function(d) d.Pages))
                For Each extra In rule.AdditionalRules.Where(Function(r) r.Active)
                    Dim hit As System.Boolean = If(extra.Kind = "exact", ContainsLiteral(allText, extra.Criteria),
                        extraction.RuleMatches.Any(Function(m) m.RuleId = extra.Id AndAlso m.Certain))
                    If hit Then
                        result.AdditionalIds.AddRange(extra.RecipientIds)
                        result.ToIds.AddRange(extra.RecipientIds)
                        result.Reasons.Add("Topic routing rule " & extra.Id & ": " & extra.Criteria)
                    End If
                Next
            End If
            If result.ToIds.Count = 0 Then
                result.ToIds.Add(rule.DefaultRecipientId)
                result.UsedDefault = True
                result.Reasons.Add("Default recipient: unresolved analysis or no named/topic recipient. Ambiguous names are never bypassed by topic rules.")
            End If
            result.ToIds = result.ToIds.Distinct().ToList()
            result.CcIds = result.CcIds.Distinct().Except(result.ToIds).ToList()
            result.CopyIds = result.CopyIds.Distinct().ToList()
            result.AdditionalIds = result.AdditionalIds.Distinct().ToList()
            result.Review.AddRange(documents.SelectMany(Function(d) d.Problems.Select(Function(p) d.FileName & ": " & p)))
            result.Review.AddRange(extraction.Problems)
            result.Review.AddRange(extraction.Sources.SelectMany(Function(s) s.Uncertainties.Select(Function(u) s.Id & ": " & u)))
            Dim zone = System.TimeZoneInfo.FindSystemTimeZoneById(rule.TimeZoneId)
            result.ReceiptDate = System.TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.SpecifyKind(receivedUtc, System.DateTimeKind.Utc), zone).Date
            result.ReceiptBasis = "Scanner mail received in " & rule.TimeZoneId & ": " & result.ReceiptDate.ToString("yyyy-MM-dd")
            ' A current stamp must belong to the selected current delivery source, never a historical enclosure.
            Dim stamps = extraction.Stamps.Where(Function(s) s.Certain AndAlso s.Relation = "current" AndAlso selected IsNot Nothing AndAlso InSource(s.Evidence, selected)).ToList()
            If stamps.Count > 0 AndAlso stamps.Select(Function(s) s.DateIso).Distinct().Count() = 1 Then
                result.UsesReceiptStamp = True
                result.ReceiptDate = IsoDate(stamps(0).DateIso).Value
                result.ReceiptBasis = "Current receipt stamp: " & Citation(stamps(0).Evidence)
            ElseIf stamps.Count > 0 Then
                result.Review.Add("Conflicting current receipt stamps; scanner mail receipt used provisionally. Receipt-dependent deadlines require review.")
            End If
            For Each stamp In extraction.Stamps.Where(Function(s) Not stamps.Contains(s))
                result.Review.Add("Receipt stamp not used (" & stamp.Relation & ", certainty=" & stamp.Certain.ToString() & "): " & Citation(stamp.Evidence))
            Next
            Dim conflicting = stamps.Select(Function(s) s.DateIso).Distinct().Count() > 1
            Dim receiptUncertain = extraction.Problems.Any(Function(p) p.StartsWith("Receipt stamp omitted", System.StringComparison.Ordinal) OrElse p.StartsWith("Optional Stamps entry omitted", System.StringComparison.Ordinal))
            For Each deadline In extraction.Deadlines
                Dim dateValue As System.DateTime? = Nothing
                Dim basis = ""
                Dim inHistoricalSource = extraction.Sources.Any(Function(s) s.FunctionInPackage = "historical" AndAlso InSource(deadline.Evidence, s))
                If deadline.Status = "current" AndAlso deadline.Certain AndAlso Not inHistoricalSource AndAlso Not System.String.IsNullOrWhiteSpace(deadline.CurrentRelation) Then
                    If deadline.Trigger = "explicit" AndAlso IsoDate(deadline.DateIso).HasValue Then
                        dateValue = IsoDate(deadline.DateIso)
                        basis = "Explicit end date."
                    ElseIf Not conflicting AndAlso Not receiptUncertain AndAlso deadline.Trigger = "receipt" AndAlso deadline.Duration > 0 AndAlso deadline.Unit = "calendar_days" AndAlso rule.DeadlineProfile = "calendar_days" AndAlso deadline.Profile = rule.DeadlineProfile Then
                        dateValue = result.ReceiptDate.AddDays(deadline.Duration)
                        basis = result.ReceiptBasis & "; day 1 is the following day; +" & deadline.Duration.ToString(System.Globalization.CultureInfo.InvariantCulture) & " calendar days."
                    End If
                End If
                Dim explicitDate = IsoDate(deadline.DateIso)
                If explicitDate.HasValue AndAlso explicitDate.Value < result.ReceiptDate Then
                    result.PastDeadlines.Add(New CalendarEvent With {.Id = "deadline-" & deadline.Id, .DateValue = explicitDate.Value, .Title = deadline.Action})
                    result.Review.Add("Deadline predates receipt; no calendar entry: " & deadline.DateIso & "; " & deadline.Action)
                ElseIf dateValue.HasValue Then
                    result.Events.Add(New CalendarEvent With {.Id = "deadline-" & deadline.Id, .DateValue = dateValue.Value,
                        .Title = deadline.Action, .Description = Citation(deadline.Evidence) & System.Environment.NewLine & basis & System.Environment.NewLine & deadline.CurrentRelation})
                Else
                    result.Review.Add("Deadline (" & deadline.Status & ", review; no calendar entry): " & deadline.Action & "; " & Citation(deadline.Evidence))
                End If
            Next
            For Each days In rule.ReminderDays
                result.Events.Add(New CalendarEvent With {.Id = "reminder-" & days.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    .DateValue = result.ReceiptDate.AddDays(days), .Title = "Wiedervorlage / Follow-up +" & days.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    .Description = result.ReceiptBasis & "; +" & days.ToString(System.Globalization.CultureInfo.InvariantCulture) & " calendar days; " & rule.TimeZoneId & If(conflicting, "; provisional receipt basis: conflicting stamps require review", "")})
            Next
            Dim today = System.TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.SpecifyKind(nowUtc, System.DateTimeKind.Utc), zone).Date
            For Each item In result.Events
                If item.DateValue < today Then result.Review.Add("Overdue / überfällig: " & item.Title & " — " & item.DateValue.ToString("yyyy-MM-dd"))
            Next
            Return result
        End Function

        Public Shared Function Citation(evidence As Evidence) As System.String
            Return evidence.DocumentId & ", p. " & evidence.StartPage.ToString(System.Globalization.CultureInfo.InvariantCulture) & "-" &
                evidence.EndPage.ToString(System.Globalization.CultureInfo.InvariantCulture) & ": " & evidence.Quote
        End Function
    End Class

    Public NotInheritable Class CalendarWriter
        Private Sub New()
        End Sub

        Private Shared Function Escape(value As System.String) As System.String
            Return If(value, "").Replace("\", "\\").Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Replace(vbLf, "\n").Replace(";", "\;").Replace(",", "\,")
        End Function

        Private Shared Sub Line(builder As System.Text.StringBuilder, value As System.String)
            ' RFC 5545: fold at 75 UTF-8 octets without splitting Unicode scalar values.
            Dim count As System.Int32 = 0
            Dim index As System.Int32 = 0
            While index < value.Length
                Dim length = If(System.Char.IsHighSurrogate(value(index)) AndAlso index + 1 < value.Length AndAlso System.Char.IsLowSurrogate(value(index + 1)), 2, 1)
                Dim piece = value.Substring(index, length)
                Dim bytes = System.Text.Encoding.UTF8.GetByteCount(piece)
                If count + bytes > 75 Then
                    builder.Append(vbCrLf).Append(" ")
                    count = 1
                End If
                builder.Append(piece)
                count += bytes
                index += length
            End While
            builder.Append(vbCrLf)
        End Sub

        Public Shared Function Render(packageId As System.String, item As CalendarEvent, createdUtc As System.DateTime) As System.String
            Dim builder As New System.Text.StringBuilder()
            For Each value In New System.String() {"BEGIN:VCALENDAR", "VERSION:2.0", "PRODID:-//Red Ink//Scan Inbox//EN", "CALSCALE:GREGORIAN", "METHOD:PUBLISH", "BEGIN:VEVENT",
                "UID:" & Registry.Hash(packageId & ":" & item.Id) & "@redink.invalid",
                "DTSTAMP:" & createdUtc.ToUniversalTime().ToString("yyyyMMdd'T'HHmmss'Z'", System.Globalization.CultureInfo.InvariantCulture),
                "DTSTART;VALUE=DATE:" & item.DateValue.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                "DTEND;VALUE=DATE:" & item.DateValue.AddDays(1).ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture),
                "SUMMARY:" & Escape(item.Title), "DESCRIPTION:" & Escape(item.Description), "TRANSP:TRANSPARENT", "END:VEVENT", "END:VCALENDAR"}
                Line(builder, value)
            Next
            Return builder.ToString()
        End Function
    End Class
End Namespace
