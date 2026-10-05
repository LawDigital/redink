' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Passive mail assessment: no browser, tooling loop, attachment saving or document opening.

' =============================================================================
' File: ThisAddIn.Commands.CheckPhishing.vb
' Purpose:
'   Passive single-mail phishing assessment with localized, validated and HTML-encoded
'   risk/confidence reports.
'
' Architecture / Function:
'   Collects mail text, literal link targets and attachment metadata without fetching
'   links or opening/executing attachments.
' =============================================================================

Option Strict On
Option Explicit On

Partial Public Class ThisAddIn
    Private _phishingCheckRunning As System.Boolean

    Public Async Sub CheckPhishing()
        If _phishingCheckRunning Then Return
        _phishingCheckRunning = True
        Try
            Dim mail As Microsoft.Office.Interop.Outlook.MailItem = Nothing
            Dim inspector = GetActiveInspector()
            If inspector IsNot Nothing Then
                mail = TryCast(inspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
            Else
                Dim explorer = Me.Application.ActiveExplorer()
                If explorer IsNot Nothing AndAlso explorer.Selection.Count > 1 Then
                    Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Select exactly one e-mail for Check Phishing.", "Red Ink - Check Phishing")
                    Return
                End If
                If explorer IsNot Nothing AndAlso explorer.Selection.Count = 1 Then
                    mail = TryCast(explorer.Selection.Item(1), Microsoft.Office.Interop.Outlook.MailItem)
                End If
            End If
            If mail Is Nothing Then
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Select or open an e-mail first.", "Red Ink - Check Phishing")
                Return
            End If
            Dim mailReference As System.String = If(mail.Subject, "(no subject)") & System.Environment.NewLine &
                If(mail.SenderName, "") & " <" & If(mail.SenderEmailAddress, "") & "> · " & mail.ReceivedTime.ToString("g", System.Globalization.CultureInfo.CurrentCulture)
            ' Snapshot Outlook COM evidence before the asynchronous model call.
            Dim evidence As New Newtonsoft.Json.Linq.JObject()
            evidence("subject") = mail.Subject
            evidence("sender_name") = mail.SenderName
            evidence("sender_address") = mail.SenderEmailAddress
            evidence("sender_address_type") = mail.SenderEmailType
            evidence("body") = mail.Body
            evidence("download_state") = mail.DownloadState.ToString()
            Dim missing As New Newtonsoft.Json.Linq.JArray()
            Try
                Try
                    evidence("transport_headers") = CStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F"))
                Catch ex As System.Exception
                    evidence("transport_headers") = CStr(mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E"))
                End Try
            Catch ex As System.Exception
                missing.Add("Transport headers unavailable: " & ex.Message)
            End Try
            Dim links As New Newtonsoft.Json.Linq.JArray()
            Try
                Dim html As New HtmlAgilityPack.HtmlDocument()
                html.LoadHtml(If(mail.HTMLBody, ""))
                Dim nodes = html.DocumentNode.SelectNodes("//*[@href or @src or @action]")
                If nodes IsNot Nothing Then
                    For Each node In nodes
                        For Each attr In New System.String() {"href", "src", "action"}
                            If node.Attributes(attr) Is Nothing Then Continue For
                            links.Add(New Newtonsoft.Json.Linq.JObject(
                                New Newtonsoft.Json.Linq.JProperty("element", node.Name),
                                New Newtonsoft.Json.Linq.JProperty("attribute", attr),
                                New Newtonsoft.Json.Linq.JProperty("target", HtmlAgilityPack.HtmlEntity.DeEntitize(node.Attributes(attr).Value)),
                                New Newtonsoft.Json.Linq.JProperty("display_text", HtmlAgilityPack.HtmlEntity.DeEntitize(node.InnerText))))
                        Next
                    Next
                End If
            Catch ex As System.Exception
                missing.Add("HTML link evidence unavailable: " & ex.Message)
            End Try
            evidence("literal_html_links_and_resources") = links
            Dim attachments As New Newtonsoft.Json.Linq.JArray()
            For index As System.Int32 = 1 To mail.Attachments.Count
                Try
                    Dim attachment = mail.Attachments.Item(index)
                    attachments.Add(New Newtonsoft.Json.Linq.JObject(
                        New Newtonsoft.Json.Linq.JProperty("filename", attachment.FileName),
                        New Newtonsoft.Json.Linq.JProperty("display_name", attachment.DisplayName),
                        New Newtonsoft.Json.Linq.JProperty("size_bytes", attachment.Size),
                        New Newtonsoft.Json.Linq.JProperty("outlook_type", attachment.Type.ToString()),
                        New Newtonsoft.Json.Linq.JProperty("inspection", "metadata only; content not opened or executed")))
                Catch ex As System.Exception
                    missing.Add("Attachment " & index.ToString() & " metadata unavailable: " & ex.Message)
                End Try
            Next
            evidence("attachments") = attachments
            evidence("missing_evidence") = missing
            Dim language As System.String = If(System.String.IsNullOrWhiteSpace(INI_Language1), "English", INI_Language1)
            Dim prompt As System.String = InterpolateAtRuntime(SP_CheckforPhishing) & System.Environment.NewLine &
                "Output language: " & language & ". Return only one JSON object, no fences, with risk (low/medium/high/undetermined), confidence (low/medium/high), risk_label and confidence_label (localized labels including the respective level), summary, confidence_reason, labels (object with summary, findings, next_steps, limitations and reference as localized section headings), findings (string array), next_steps (string array), limitations (string). All display text must use the output language. Explain missing evidence explicitly. Use concise findings and practical next steps; state when no warning signs were found. Keep the displayed report below 9000 characters."
            Dim raw As System.String = Await LLM(prompt, evidence.ToString(Newtonsoft.Json.Formatting.None), HideSplash:=False, ToolExecution:=False)
            Dim result As Newtonsoft.Json.Linq.JObject = Nothing
            Dim invalid As System.Exception = Nothing
            Try
                result = ParsePhishingAssessment(raw)
            Catch ex As System.Exception
                invalid = ex
            End Try
            If invalid IsNot Nothing Then
                System.Diagnostics.Debug.WriteLine("[CheckPhishing] Assessment format rejected; retrying once: " & invalid.Message)
                raw = Await LLM(prompt & System.Environment.NewLine & "Your previous response did not meet the JSON contract. Return a complete valid JSON object only; no introduction or Markdown.", evidence.ToString(Newtonsoft.Json.Formatting.None), HideSplash:=False, ToolExecution:=False)
                result = ParsePhishingAssessment(raw)
            End If
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowHTMLCustomMessageBox(RenderPhishingAssessment(result, mailReference), "Red Ink - Check Phishing")
        Catch ex As System.Exception
            Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Phishing assessment could not be completed: " & ex.Message, "Red Ink - Check Phishing")
        Finally
            _phishingCheckRunning = False
        End Try
    End Sub
    Private Shared Function ParsePhishingAssessment(raw As System.String) As Newtonsoft.Json.Linq.JObject
        Dim text = If(raw, "").Trim()
        If text.StartsWith("""", System.StringComparison.Ordinal) Then
            Dim decoded = Newtonsoft.Json.Linq.JToken.Parse(text)
            If decoded.Type = Newtonsoft.Json.Linq.JTokenType.String Then text = CStr(decoded).Trim()
        End If
        ' Accept Markdown fences and surrounding prose, but require one complete validated object.
        Dim first = text.IndexOf("{"c), last = text.LastIndexOf("}"c)
        If first < 0 OrElse last <= first Then Throw New System.IO.InvalidDataException("The AI did not return a structured assessment.")
        Dim result = Newtonsoft.Json.Linq.JObject.Parse(text.Substring(first, last - first + 1))
        For Each field In New System.String() {"risk_label", "confidence_label", "summary", "confidence_reason", "limitations"}
            If result(field) Is Nothing OrElse result(field).Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse System.String.IsNullOrWhiteSpace(CStr(result(field))) Then Throw New System.IO.InvalidDataException("Incomplete assessment: " & field)
        Next
        For Each field In New System.String() {"risk", "confidence"}
            If result(field) Is Nothing OrElse result(field).Type <> Newtonsoft.Json.Linq.JTokenType.String Then Throw New System.IO.InvalidDataException("Missing assessment level.")
            result(field) = CStr(result(field)).Trim().ToLowerInvariant()
        Next
        Dim riskLevels As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {"low", "medium", "high", "undetermined"}
        Dim confidenceLevels As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.Ordinal) From {"low", "medium", "high"}
        If Not riskLevels.Contains(CStr(result("risk"))) OrElse Not confidenceLevels.Contains(CStr(result("confidence"))) Then Throw New System.IO.InvalidDataException("Invalid risk or confidence level.")
        For Each field In New System.String() {"findings", "next_steps"}
            Dim values = TryCast(result(field), Newtonsoft.Json.Linq.JArray)
            If values Is Nothing OrElse values.Count = 0 Then Throw New System.IO.InvalidDataException("Missing assessment section: " & field)
            For Each value In values
                If value.Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse System.String.IsNullOrWhiteSpace(CStr(value)) Then Throw New System.IO.InvalidDataException("Invalid assessment item.")
            Next
        Next
        Dim labels = TryCast(result("labels"), Newtonsoft.Json.Linq.JObject)
        If labels Is Nothing Then Throw New System.IO.InvalidDataException("Missing localized headings.")
        For Each field In New System.String() {"summary", "findings", "next_steps", "limitations", "reference"}
            If labels(field) Is Nothing OrElse labels(field).Type <> Newtonsoft.Json.Linq.JTokenType.String OrElse System.String.IsNullOrWhiteSpace(CStr(labels(field))) Then Throw New System.IO.InvalidDataException("Missing localized heading: " & field)
        Next
        If result.ToString(Newtonsoft.Json.Formatting.None).Length > 14000 Then Throw New System.IO.InvalidDataException("Assessment exceeds the display limit; no truncated report is accepted.")
        Return result
    End Function

    Private Shared Function PhishingHtmlText(value As System.String) As System.String
        Return System.Net.WebUtility.HtmlEncode(If(value, "")).Replace(System.Environment.NewLine, "<br>").Replace(vbLf, "<br>")
    End Function

    Private Shared Function RenderPhishingAssessment(result As Newtonsoft.Json.Linq.JObject, mailReference As System.String) As System.String
        Dim risk = CStr(result("risk"))
        Dim color As System.String = "#64748b", symbol As System.String = "?"
        Select Case risk
            Case "low" : color = "#167447" : symbol = "&#10003;"
            Case "medium" : color = "#9a6400" : symbol = "&#9888;"
            Case "high" : color = "#b42318" : symbol = "!"
        End Select
        Dim labels = CType(result("labels"), Newtonsoft.Json.Linq.JObject)
        Dim html As New System.Text.StringBuilder("<!doctype html><html><head><meta http-equiv='X-UA-Compatible' content='IE=edge'><meta charset='utf-8'><style>body{font-family:Segoe UI,Arial,sans-serif;background:#f2f5f9;color:#203047;margin:0;padding:26px}h1{font-size:23px;margin:0}h2{font-size:16px;margin:0 0 10px}.card{background:#fff;border:1px solid #dde4ee;border-radius:10px;padding:20px;margin-bottom:16px}.muted{color:#526477;font-size:13px}p{line-height:1.6;margin:8px 0}li{line-height:1.6;margin:7px 0}.badge{display:inline-block;padding:6px 12px;border-radius:20px;background:#edf2f7;margin-top:12px}.hero{border-left:7px solid ")
        html.Append(color).Append("}.symbol{font-size:32px;font-weight:bold;color:").Append(color).Append(";margin-right:12px}</style></head><body>")
        html.Append("<div class='card muted'><b>&#9993; ").Append(PhishingHtmlText(CStr(labels("reference")))).Append("</b><p>").Append(PhishingHtmlText(mailReference)).Append("</p></div>")
        html.Append("<div class='card hero'><h1><span class='symbol'>").Append(symbol).Append("</span>").Append(PhishingHtmlText(CStr(result("risk_label")))).Append("</h1><p>").Append(PhishingHtmlText(CStr(result("summary")))).Append("</p><div class='badge'>&#9678; ").Append(PhishingHtmlText(CStr(result("confidence_label")))).Append("</div><p class='muted'>").Append(PhishingHtmlText(CStr(result("confidence_reason")))).Append("</p></div>")
        For Each field In New System.String() {"findings", "next_steps"}
            html.Append("<div class='card'><h2>").Append(If(field = "findings", "&#128269; ", "&#10140; ")).Append(PhishingHtmlText(CStr(labels(field)))).Append("</h2><ul>")
            For Each value In CType(result(field), Newtonsoft.Json.Linq.JArray)
                html.Append("<li>").Append(PhishingHtmlText(CStr(value))).Append("</li>")
            Next
            html.Append("</ul></div>")
        Next
        html.Append("<div class='card muted'><h2>&#9432; ").Append(PhishingHtmlText(CStr(labels("limitations")))).Append("</h2><p>").Append(PhishingHtmlText(CStr(result("limitations")))).Append("</p></div></body></html>")
        ' All mail/model content is encoded text. No links, images, scripts or external styles are generated.
        Return html.ToString()
    End Function

End Class
