' Ask Inky private per-user storage. All generated files live under this one root.
Friend NotInheritable Class StoragePaths
    Private Sub New()
    End Sub
    Friend Shared ReadOnly Root As String = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "AskInky")
End Class
