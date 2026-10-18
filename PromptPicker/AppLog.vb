Friend NotInheritable Class AppLog
    Private Sub New()
    End Sub

    Public Shared ReadOnly Property Root As String
        Get
            Return System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData), "Red Ink", "PromptPicker")
        End Get
    End Property

    Public Shared Sub Write(message As String, Optional ex As System.Exception = Nothing)
        Try
            System.IO.Directory.CreateDirectory(Root)
            Dim entry As String = System.DateTimeOffset.Now.ToString("O", System.Globalization.CultureInfo.InvariantCulture) & " " & message & If(ex Is Nothing, "", " " & ex.ToString()) & System.Environment.NewLine
            System.IO.File.AppendAllText(System.IO.Path.Combine(Root, "promptpicker.log"), entry)
        Catch
            ' Diagnostics must never prevent the tray process from running.
        End Try
    End Sub
End Class
