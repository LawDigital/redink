' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: UiStrings.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Original Office Web settings translations plus desktop lifecycle strings.
Friend NotInheritable Class UiStrings
    Private Shared ReadOnly Bundles As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))(System.StringComparer.Ordinal)

    Public Shared Function NormalizeLanguage(value As String) As String
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "en", "e" : Return "en"
            Case "fr", "f" : Return "fr"
            Case "it" : Return "it"
            Case Else : Return "de"
        End Select
    End Function

    Public Shared Function IsSupported(value As String) As Boolean
        Select Case If(value, "").Trim().ToLowerInvariant()
            Case "de", "d", "en", "e", "fr", "f", "it" : Return True
            Case Else : Return False
        End Select
    End Function

    Public Shared Function GetBundle(language As String) As System.Collections.Generic.Dictionary(Of String, String)
        Dim code As String = NormalizeLanguage(language)
        SyncLock Bundles
            Dim result As System.Collections.Generic.Dictionary(Of String, String) = Nothing
            If Not Bundles.TryGetValue(code, result) Then
                result = System.Text.Json.JsonSerializer.Deserialize(Of System.Collections.Generic.Dictionary(Of String, String))(WebAssets.ReadText("I18n/" & code & ".json"))
                If result Is Nothing Then Throw New System.IO.InvalidDataException("Translation resource is empty")
                Bundles.Add(code, result)
            End If
            Return result
        End SyncLock
    End Function

    Public Shared Function GetAllBundles() As System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))
        Dim result As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.Dictionary(Of String, String))()
        For Each code As String In New String() {"de", "en", "fr", "it"}
            result.Add(code, GetBundle(code))
        Next
        Return result
    End Function

    Public Shared Function GetText(key As String, language As String) As String
        Dim value As String = Nothing
        If GetBundle(language).TryGetValue(key, value) OrElse GetBundle("en").TryGetValue(key, value) Then Return value
        ClientLog.Write("translation.missing." & key)
        Return key
    End Function
End Class
