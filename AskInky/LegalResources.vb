' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: LegalResources.vb
' Purpose: Provides the unmodified bundled Red Ink license and third-party notices.
' Architecture: Loads text embedded by AskInky.vbproj for portable single-EXE use.
' =============================================================================
Option Strict On
Option Explicit On

Friend NotInheritable Class LegalResources
    Private Sub New()
    End Sub
    Public Shared ReadOnly Property LicenseText As String
        Get
            Using stream As System.IO.Stream = GetType(LegalResources).Assembly.GetManifestResourceStream("AskInky.LICENSE.txt")
                If stream Is Nothing Then Throw New System.IO.FileNotFoundException("Embedded license missing")
                Using reader As New System.IO.StreamReader(stream, System.Text.Encoding.UTF8, True)
                    Return reader.ReadToEnd()
                End Using
            End Using
        End Get
    End Property
End Class
