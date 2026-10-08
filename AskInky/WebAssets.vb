' Part of Red Ink (AskInky). Embedded, offline web assets; never loads user-selected files.
Friend NotInheritable Class WebAssets
    Private Sub New()
    End Sub

    Public Shared Function ReadBytes(name As String) As Byte()
        Using source As System.IO.Stream = GetType(WebAssets).Assembly.GetManifestResourceStream("AskInky.Web." & name.Replace("/", "."))
            If source Is Nothing Then Throw New System.IO.FileNotFoundException("Embedded UI resource missing", name)
            Using buffer As New System.IO.MemoryStream()
                source.CopyTo(buffer)
                Return buffer.ToArray()
            End Using
        End Using
    End Function

    Public Shared Function ReadText(name As String) As String
        Return System.Text.Encoding.UTF8.GetString(ReadBytes(name))
    End Function

    Public Shared Function IsPublicAsset(name As String) As Boolean
        Select Case name
            Case "assistant.html", "settings.html", "office-web.css", "desktop.css", "common.js", "rendering.js", "assistant.js", "settings.js"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Public Shared Function ContentType(name As String) As String
        Select Case System.IO.Path.GetExtension(name)
            Case ".html" : Return "text/html; charset=utf-8"
            Case ".css" : Return "text/css; charset=utf-8"
            Case ".js" : Return "text/javascript; charset=utf-8"
            Case Else : Return "application/octet-stream"
        End Select
    End Function
End Class
