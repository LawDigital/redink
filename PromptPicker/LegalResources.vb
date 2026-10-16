' Part of Red Ink (Prompt Picker). Embedded license and notices, offline only.
Option Strict On
Option Explicit On

Friend NotInheritable Class LegalResources
    Private Sub New()
    End Sub

    Public Shared ReadOnly Property LicenseText As String
        Get
            Using stream As System.IO.Stream = GetType(LegalResources).Assembly.GetManifestResourceStream("PromptPicker.LICENSE.txt")
                If stream Is Nothing Then Throw New System.IO.FileNotFoundException("Eingebettete Lizenzhinweise fehlen.")
                Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8, True)
                    Return reader.ReadToEnd()
                End Using
            End Using
        End Get
    End Property
End Class
