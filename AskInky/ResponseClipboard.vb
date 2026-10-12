' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ResponseClipboard.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). One Markdown source; copy the same sanitized DOM that is displayed.
Imports Markdig

Friend NotInheritable Class ResponseClipboard
    Private Shared ReadOnly Pipeline As Markdig.MarkdownPipeline = New Markdig.MarkdownPipelineBuilder().UseAdvancedExtensions().DisableHtml().Build()

    Public Shared Function Render(markdown As String) As String
        Return Markdig.Markdown.ToHtml(If(markdown, ""), Pipeline)
    End Function

    Friend Shared Function BuildHtmlFormat(fragment As String) As String
        Dim prefix As String = "<html><head><meta charset=""utf-8""></head><body><!--StartFragment-->"
        Dim suffix As String = "<!--EndFragment--></body></html>"
        Dim template As String = "Version:1.0" & vbCrLf & "StartHTML:{0:0000000000}" & vbCrLf & "EndHTML:{1:0000000000}" & vbCrLf & "StartFragment:{2:0000000000}" & vbCrLf & "EndFragment:{3:0000000000}" & vbCrLf
        Dim culture As System.Globalization.CultureInfo = System.Globalization.CultureInfo.InvariantCulture
        Dim startHtml As Integer = System.Text.Encoding.UTF8.GetByteCount(String.Format(culture, template, 0, 0, 0, 0))
        Dim startFragment As Integer = startHtml + System.Text.Encoding.UTF8.GetByteCount(prefix)
        Dim endFragment As Integer = startFragment + System.Text.Encoding.UTF8.GetByteCount(fragment)
        Dim endHtml As Integer = endFragment + System.Text.Encoding.UTF8.GetByteCount(suffix)
        Return String.Format(culture, template, startHtml, endHtml, startFragment, endFragment) & prefix & fragment & suffix
    End Function

    Public Shared Sub CopyRendered(fragment As String, plainText As String)
        If fragment Is Nothing OrElse plainText Is Nothing Then Throw New System.ArgumentNullException(NameOf(fragment))
        Dim data As New System.Windows.Forms.DataObject()
        ' CF_HTML is UTF-8 with byte offsets; UnicodeText is the native UTF-16 plain-text format.
        data.SetData(System.Windows.Forms.DataFormats.Html, False, BuildHtmlFormat(fragment))
        data.SetData(System.Windows.Forms.DataFormats.UnicodeText, False, plainText)
        System.Windows.Forms.Clipboard.SetDataObject(data, True, 5, 80)
    End Sub
End Class
