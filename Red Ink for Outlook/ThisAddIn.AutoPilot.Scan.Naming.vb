' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Naming.vb
' Purpose: Semantic naming suggestions; deterministic, portable filenames.
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
    Public Class NamingResponse
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Documents As New System.Collections.Generic.List(Of NamingSuggestion)()
    End Class

    Public Class NamingSuggestion
        Public Property DocumentId As System.String = ""
        Public Property Content As System.String = ""
        ' References an extracted, evidenced letter date; never a free model date.
        Public Property DateSourceId As System.String = ""
    End Class

    Public Class AttachmentName
        Public Property DocumentId As System.String = ""
        Public Property OriginalFileName As System.String = ""
        Public Property FileName As System.String = ""
        Public Property Content As System.String = ""
        Public Property DateSourceId As System.String = ""
        Public Property Note As System.String = ""
    End Class

    Public NotInheritable Class AttachmentNaming
        Public Const DefaultTemplate As System.String = "{received:yyMMdd} {content}"
        Private Const MaximumFileNameLength As System.Int32 = 140
        Private Shared ReadOnly Tokens As New System.Text.RegularExpressions.Regex("\{(?<name>received|documentDate|index|package|content|original)(?::(?<format>[^{}]+))?\}", System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        Private Sub New()
        End Sub

        Public Shared Function IsStandardTemplate(template As System.String) As System.Boolean
            Return template = DefaultTemplate OrElse template = "{received:yyyyMMdd_HHmmss}_{index:000}_{content}" OrElse template = "{received:yyyyMMdd_HHmmss}_{package}_{index:000}_{content}"
        End Function

        Public Shared Sub ValidateTemplate(template As System.String)
            If System.String.IsNullOrWhiteSpace(template) OrElse template.Length > 500 Then Throw New System.IO.InvalidDataException("AttachmentNameTemplate must contain 1..500 characters.")
            Dim remainder = Tokens.Replace(template, "")
            If remainder.Contains("{") OrElse remainder.Contains("}") Then Throw New System.IO.InvalidDataException("Unknown or malformed attachment filename token.")
            For Each hit As System.Text.RegularExpressions.Match In Tokens.Matches(template)
                Dim token = hit.Groups("name").Value
                Dim format = hit.Groups("format").Value
                If format.Length = 0 Then Continue For
                Select Case token
                    Case "received"
                        System.DateTime.UtcNow.ToString(format, System.Globalization.CultureInfo.InvariantCulture)
                    Case "documentDate"
                        If System.Text.RegularExpressions.Regex.IsMatch(format, "[^yMd._\-/ ]") Then Throw New System.IO.InvalidDataException("documentDate permits date components y/M/d and separators only; a letter time is not known.")
                        System.DateTime.UtcNow.ToString(format, System.Globalization.CultureInfo.InvariantCulture)
                    Case "index"
                        If Not System.Text.RegularExpressions.Regex.IsMatch(format, "\A(?:0{1,8}|[dD][1-8]?)\z") Then Throw New System.IO.InvalidDataException("index format must be 0..00000000 or D/D1..D8.")
                    Case Else
                        Throw New System.IO.InvalidDataException("This attachment filename token has no format argument: " & token)
                End Select
            Next
        End Sub

        Public Shared Function SafeStem(value As System.String) As System.String
            Dim normalized As System.String
            Try
                normalized = If(value, "").Normalize(System.Text.NormalizationForm.FormKC)
            Catch ex As System.ArgumentException
                normalized = ""
            End Try
            Dim builder As New System.Text.StringBuilder()
            For Each character In normalized
                If System.Char.IsControl(character) OrElse "<>:""/\|?*".IndexOf(character) >= 0 OrElse
                   System.Globalization.CharUnicodeInfo.GetUnicodeCategory(character) = System.Globalization.UnicodeCategory.Format Then
                    builder.Append("_"c)
                Else
                    builder.Append(character)
                End If
            Next
            Dim result = System.Text.RegularExpressions.Regex.Replace(builder.ToString(), "\s+", "_").Trim(" "c, "."c, "_"c)
            If result = "" Then result = "000"
            Dim deviceStem = result.Split("."c)(0)
            If System.Text.RegularExpressions.Regex.IsMatch(deviceStem, "\A(?:CON|PRN|AUX|NUL|COM[0-9]|LPT[0-9])\z", System.Text.RegularExpressions.RegexOptions.IgnoreCase Or System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then result = "_" & result
            Return result
        End Function

        Private Shared Function TrimLength(value As System.String, maximum As System.Int32, Optional wholeWords As System.Boolean = False) As System.String
            If value.Length <= maximum Then Return value
            Dim length = maximum
            If length > 0 AndAlso System.Char.IsHighSurrogate(value(length - 1)) Then length -= 1
            If Not wholeWords Then Return value.Substring(0, length).TrimEnd(" "c, "."c)
            ' Keep complete words/components. A token without any boundary is not cut in half.
            Dim boundary = value.LastIndexOfAny(New System.Char() {" "c, "_"c, "-"c}, length - 1, length)
            If boundary > 0 Then Return value.Substring(0, boundary).TrimEnd(" "c, "."c, "_"c, "-"c)
            Return "000"
        End Function

        Public Shared Function Create(rule As RuleSet, operation As Operation, suggestions As NamingResponse,
                                      Optional fallbackReason As System.String = "", Optional fileNameLimit As System.Int32 = MaximumFileNameLength) As System.Collections.Generic.List(Of AttachmentName)
            ValidateTemplate(rule.AttachmentNameTemplate)
            fileNameLimit = System.Math.Min(fileNameLimit, MaximumFileNameLength)
            If fileNameLimit < 32 Then Throw New System.IO.PathTooLongException("The scan staging/retention paths leave insufficient room for a meaningful portable attachment filename.")
            If suggestions Is Nothing OrElse suggestions.Documents Is Nothing Then Throw New System.IO.InvalidDataException("Missing attachment naming suggestions.")
            If suggestions.Documents.Any(Function(s) s Is Nothing) OrElse suggestions.Documents.Select(Function(s) s.DocumentId).Distinct().Count() <> suggestions.Documents.Count Then Throw New System.IO.InvalidDataException("Duplicate/null attachment naming suggestions.")
            If suggestions.Documents.Any(Function(s) s.Content IsNot Nothing AndAlso s.Content.Length > 2000) Then Throw New System.IO.InvalidDataException("Naming content label is too long.")
            If suggestions.Documents.Any(Function(s) Not operation.Documents.Any(Function(d) d.Id = s.DocumentId)) Then Throw New System.IO.InvalidDataException("Unknown naming document ID.")
            If fallbackReason = "" AndAlso (suggestions.Documents.Count <> operation.Documents.Count OrElse suggestions.Documents.Any(Function(s) System.String.IsNullOrWhiteSpace(s.Content))) Then Throw New System.IO.InvalidDataException("Naming must cover every attachment with a content label.")
            Dim used As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase) From {"scan-package.json", "meta.json", "operation.json", "processing.lock"}
            If operation.Decision IsNot Nothing Then
                Dim calendarNumber As System.Int32 = 0
                For Each item In operation.Decision.Events
                    calendarNumber += 1
                    used.Add(item.DateValue.ToString("yyyyMMdd") & "_" & calendarNumber.ToString("00") & ".ics")
                Next
            End If
            Dim received = System.TimeZoneInfo.ConvertTimeFromUtc(System.DateTime.SpecifyKind(operation.ReceivedUtc, System.DateTimeKind.Utc), System.TimeZoneInfo.FindSystemTimeZoneById(rule.TimeZoneId))
            Dim result As New System.Collections.Generic.List(Of AttachmentName)()
            For index As System.Int32 = 0 To operation.Documents.Count - 1
                Dim document = operation.Documents(index)
                Dim suggestion = suggestions.Documents.SingleOrDefault(Function(s) s.DocumentId = document.Id)
                Dim original = System.IO.Path.GetFileNameWithoutExtension(document.FileName)
                Dim content = If(suggestion Is Nothing OrElse System.String.IsNullOrWhiteSpace(suggestion.Content), original, suggestion.Content)
                Dim dateSource As Source = Nothing
                If suggestion IsNot Nothing AndAlso Not System.String.IsNullOrWhiteSpace(suggestion.DateSourceId) Then
                    dateSource = operation.Extraction.Sources.SingleOrDefault(Function(s) s.Id = suggestion.DateSourceId AndAlso s.Certain AndAlso s.Evidence.DocumentId = document.Id AndAlso Decisions.IsoDate(s.DateIso).HasValue)
                    If dateSource Is Nothing Then Throw New System.IO.InvalidDataException("Naming date does not reference a verified letter date in the same attachment.")
                End If
                Dim ordinal = index + 1
                Dim standard = IsStandardTemplate(rule.AttachmentNameTemplate)
                Dim effectiveTemplate = If(standard, DefaultTemplate, rule.AttachmentNameTemplate)
                Dim stem = Tokens.Replace(effectiveTemplate,
                    Function(hit As System.Text.RegularExpressions.Match) As System.String
                        Dim format = hit.Groups("format").Value
                        Select Case hit.Groups("name").Value
                            Case "received" : Return received.ToString(If(format = "", "yyyyMMdd_HHmmss", format), System.Globalization.CultureInfo.InvariantCulture)
                            Case "documentDate" : Return If(dateSource Is Nothing, "00000000", Decisions.IsoDate(dateSource.DateIso).Value.ToString(If(format = "", "yyyyMMdd", format), System.Globalization.CultureInfo.InvariantCulture))
                            Case "index" : Return ordinal.ToString(If(format = "", "000", format), System.Globalization.CultureInfo.InvariantCulture)
                            Case "package" : Return Registry.Hash(operation.Id).Substring(0, 10)
                            Case "content" : Return SafeStem(content)
                            Case "original" : Return SafeStem(original)
                            Case Else : Throw New System.IO.InvalidDataException("Unknown filename token.")
                        End Select
                    End Function)
                Dim extension = System.IO.Path.GetExtension(document.FileName)
                If extension.Length > 20 OrElse extension.Any(Function(c) System.Char.IsControl(c) OrElse "<>:""/\|?*".IndexOf(c) >= 0) Then Throw New System.IO.InvalidDataException("Unsafe original attachment extension.")
                stem = SafeStem(stem)
                If standard Then stem = System.Text.RegularExpressions.Regex.Replace(stem.Replace("_", " "), "\s+", " ").Trim()
                Dim suffix As System.String = ""
                Dim attempt As System.Int32 = 1
                Dim name As System.String
                Do
                    Dim shortened = TrimLength(stem, fileNameLimit - extension.Length - suffix.Length, standard)
                    If standard AndAlso (shortened = "000" OrElse System.Text.RegularExpressions.Regex.IsMatch(shortened, "\A[0-9]{6}\z")) Then Throw New System.IO.InvalidDataException("No complete descriptive word fits the attachment filename budget; choose a shorter naming instruction or storage path.")
                    name = shortened & suffix & extension
                    If used.Add(name) Then Exit Do
                    attempt += 1
                    suffix = If(standard, " (" & attempt.ToString(System.Globalization.CultureInfo.InvariantCulture) & ")", "_" & attempt.ToString("000", System.Globalization.CultureInfo.InvariantCulture))
                Loop
                result.Add(New AttachmentName With {.DocumentId = document.Id, .OriginalFileName = document.FileName, .FileName = name,
                    .Content = content, .DateSourceId = If(dateSource Is Nothing, "", dateSource.Id),
                    .Note = fallbackReason & If(stem.Length + extension.Length + suffix.Length > fileNameLimit, "; filename shortened to fit the portable path limit", "") & If(dateSource Is Nothing AndAlso rule.AttachmentNameTemplate.Contains("{documentDate"), "; no evidenced letter date: 00000000", "")})
            Next
            Return result
        End Function
    End Class
End Namespace
