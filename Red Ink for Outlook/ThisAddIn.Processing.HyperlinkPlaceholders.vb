' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Processing.HyperlinkPlaceholders.vb
' Purpose: Round-trips hyperlinks inside an Outlook (Word editor) selection
'          through an LLM / ReviewChangesDialog pipeline by encoding them as
'          standard Markdown link syntax:  [display](url)
'
'          This format is preserved naturally by the LLM, renders nicely in
'          the review dialog, and is turned back into a real hyperlink by
'          SLib.InsertTextWithMarkdown via Markdig. No post-insert restore
'          step is required for the Markdown insertion path.
'
' KEY ARCHITECTURE
'   • Surface: ThisAddIn (host partial/UI).
'   • Entry points/helpers: EncodeHyperlinksAsMarkdown, ConvertAnchorsToMarkdown,
'     EscapeMarkdownLinkLabel.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Text
Imports System.Text.RegularExpressions
Imports HtmlAgilityPack
Imports Microsoft.Office.Interop.Word

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Replaces every hyperlink inside <paramref name="range"/> with a Markdown
    ''' link of the form "[display](url)". The original Word hyperlink field is
    ''' removed; the resulting plain text round-trips losslessly through the LLM
    ''' and the review dialog, and is reconstituted as a real hyperlink by
    ''' SLib.InsertTextWithMarkdown when the reviewed text is re-inserted.
    ''' </summary>
    Public Shared Sub EncodeHyperlinksAsMarkdown(ByVal range As Range)
        If range Is Nothing Then Return

        Try
            ' Snapshot in reverse document order so position shifts don't break us.
            Dim links As New List(Of Hyperlink)()
            For Each h As Hyperlink In range.Hyperlinks
                links.Add(h)
            Next
            links.Sort(Function(a, b) b.Range.Start.CompareTo(a.Range.Start))

            For Each h As Hyperlink In links
                Dim hr As Range = h.Range
                Dim url As String = If(h.Address, String.Empty)
                If Not String.IsNullOrEmpty(h.SubAddress) Then
                    url &= "#" & h.SubAddress
                End If
                Dim disp As String = If(hr.Text, String.Empty)
                If String.IsNullOrEmpty(disp) Then disp = url
                If String.IsNullOrEmpty(url) Then Continue For

                hr.Text = "[" & EscapeMarkdownLinkLabel(disp) & "](" & EscapeMarkdownLinkUrl(url) & ")"
            Next
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("EncodeHyperlinksAsMarkdown: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Rewrites every &lt;a href="…"&gt;label&lt;/a&gt; element in the supplied
    ''' HTML as plain "[label](url)" Markdown so that downstream HTML stripping
    ''' (e.g. SLib.RemoveHTML for the review dialog) does not lose the URL.
    ''' </summary>
    Public Shared Function ConvertAnchorsToMarkdown(ByVal html As String) As String
        If String.IsNullOrEmpty(html) Then Return html

        Try
            Dim doc As New HtmlDocument()
            doc.LoadHtml(html)

            Dim anchors = doc.DocumentNode.SelectNodes("//a[@href]")
            If anchors Is Nothing Then Return html

            For Each a As HtmlNode In anchors
                Dim url As String = HtmlEntity.DeEntitize(a.GetAttributeValue("href", String.Empty))
                Dim disp As String = HtmlEntity.DeEntitize(a.InnerText)
                If String.IsNullOrEmpty(disp) Then disp = url
                If String.IsNullOrEmpty(url) Then Continue For

                Dim replacement As String = "[" & EscapeMarkdownLinkLabel(disp) & "](" & EscapeMarkdownLinkUrl(url) & ")"
                Dim textNode As HtmlNode = doc.CreateTextNode(HtmlEntity.Entitize(replacement))
                a.ParentNode.ReplaceChild(textNode, a)
            Next

            Return doc.DocumentNode.OuterHtml
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("ConvertAnchorsToMarkdown: " & ex.Message)
            Return html
        End Try
    End Function

    Private Shared Function EscapeMarkdownLinkLabel(value As String) As String
        Dim s As String = If(value, String.Empty)
        ' Avoid breaking the [..](..) structure
        s = s.Replace("\", "\\")
        s = s.Replace("[", "\[")
        s = s.Replace("]", "\]")
        s = s.Replace(vbCr, " ").Replace(vbLf, " ")
        Return s.Trim()
    End Function

    Private Shared Function EscapeMarkdownLinkUrl(value As String) As String
        Dim s As String = If(value, String.Empty).Trim()
        s = s.Replace(" ", "%20")
        s = s.Replace("(", "%28")
        s = s.Replace(")", "%29")
        Return s
    End Function

End Class