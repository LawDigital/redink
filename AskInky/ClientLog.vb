' Part of Red Ink (AskInky). Diagnostics deliberately exclude text, URLs and credentials.
Friend NotInheritable Class ClientLog
    Private Sub New()
    End Sub

    Public Shared Sub Write(operation As String, Optional failure As System.Exception = Nothing)
        Try
            Dim folder As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RedInk", "AskInky")
            System.IO.Directory.CreateDirectory(folder)
            Dim path As String = System.IO.Path.Combine(folder, "diagnostics.log")
            If System.IO.File.Exists(path) AndAlso New System.IO.FileInfo(path).Length > 1048576 Then
                System.IO.File.Move(path, path & ".previous", True)
            End If
            Dim detail As String = If(failure Is Nothing, "", " | " & failure.GetType().FullName & " | HResult=" & failure.HResult.ToString(System.Globalization.CultureInfo.InvariantCulture))
            System.IO.File.AppendAllText(path, System.DateTimeOffset.UtcNow.ToString("O", System.Globalization.CultureInfo.InvariantCulture) & " | " & operation & detail & System.Environment.NewLine, New System.Text.UTF8Encoding(False))
        Catch logError As System.Exception
            System.Diagnostics.Debug.WriteLine("Diagnostic write failed: " & logError.GetType().FullName)
        End Try
    End Sub
End Class
