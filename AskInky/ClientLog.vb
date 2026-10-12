' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: ClientLog.vb
' Purpose: Ask Inky desktop functionality in the modular Windows client.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). Diagnostics deliberately exclude text, URLs and credentials.
Friend NotInheritable Class ClientLog
    Private Sub New()
    End Sub

    Public Shared Sub Write(operation As String, Optional failure As System.Exception = Nothing)
        Try
            Dim folder As String = StoragePaths.Root
            System.IO.Directory.CreateDirectory(folder)
            Dim path As String = System.IO.Path.Combine(folder, "diagnostics.log")
            If System.IO.File.Exists(path) AndAlso New System.IO.FileInfo(path).Length > 1048576 Then
                System.IO.File.Move(path, path & ".previous", True)
            End If
            Dim detail As String = If(failure Is Nothing, "", " | " & failure.GetType().FullName & " | HResult=" & failure.HResult.ToString(System.Globalization.CultureInfo.InvariantCulture))
            If failure IsNot Nothing Then
                ' HelperFailure metadata is part of the fixed client error contract, not a server response.
                ' Never log Exception.Message, URLs, response bodies or credentials.
                Dim helper As HelperFailure = TryCast(failure, HelperFailure)
                If helper IsNot Nothing Then
                    Dim key As String = If(helper.ResourceKey, "")
                    If key.Length > 0 AndAlso key.Length <= 96 AndAlso System.Text.RegularExpressions.Regex.IsMatch(key, "\A[a-zA-Z0-9_.-]+\z", System.Text.RegularExpressions.RegexOptions.CultureInvariant) Then
                        detail &= " | ResourceKey=" & key
                    End If
                    detail &= " | HttpStatus=" & helper.HttpStatus.ToString(System.Globalization.CultureInfo.InvariantCulture)
                End If
                ' Diagnostic call sites without messages, argument values, user data or source paths.
                Dim methods As New System.Collections.Generic.List(Of String)()
                Dim frames As System.Diagnostics.StackFrame() = New System.Diagnostics.StackTrace(failure, False).GetFrames()
                If frames IsNot Nothing Then
                    For Each frame As System.Diagnostics.StackFrame In frames
                        Dim method As System.Reflection.MethodBase = frame.GetMethod()
                        If method IsNot Nothing AndAlso method.DeclaringType IsNot Nothing Then methods.Add(method.DeclaringType.FullName & "." & method.Name)
                        If methods.Count >= 8 Then Exit For
                    Next
                End If
                If methods.Count > 0 Then detail &= " | frames=" & String.Join(" <- ", methods)
            End If
            System.IO.File.AppendAllText(path, System.DateTimeOffset.UtcNow.ToString("O", System.Globalization.CultureInfo.InvariantCulture) & " | " & operation & detail & System.Environment.NewLine, New System.Text.UTF8Encoding(False))
        Catch logError As System.Exception
            System.Diagnostics.Debug.WriteLine("Diagnostic write failed: " & logError.GetType().FullName)
        End Try
    End Sub
End Class
