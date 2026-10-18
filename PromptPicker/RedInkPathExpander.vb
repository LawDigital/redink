' Standalone equivalent of SharedMethods.VarHelpers.vb:ExpandEnvironmentVariables.
' Do not change the Word path semantics when consuming PromptLib / PromptLibLocal.
Friend NotInheritable Class RedInkPathExpander
    Private Sub New()
    End Sub

    Public Shared Function Expand(filePath As String) As String
        If System.String.IsNullOrWhiteSpace(filePath) Then Return ""
        Dim expanded As String = System.Environment.ExpandEnvironmentVariables(filePath.Trim())
        Try
            expanded = expanded.Trim(""""c)
            expanded = ReplaceVariable(expanded, "%APPDATA%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.ApplicationData))
            expanded = ReplaceVariable(expanded, "%LOCALAPPDATA%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData))
            expanded = ReplaceVariable(expanded, "%PROGRAMDATA%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData))
            expanded = ReplaceVariable(expanded, "%USERPROFILE%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile))
            expanded = ReplaceVariable(expanded, "%WINDIR%", System.Environment.GetEnvironmentVariable("WINDIR"))
            expanded = ReplaceVariable(expanded, "%TEMP%", System.IO.Path.GetTempPath())
            expanded = ReplaceVariable(expanded, "%HOMEPATH%", System.Environment.GetEnvironmentVariable("HOMEPATH"))
            expanded = ReplaceVariable(expanded, "%APPSTARTUPPATH%", System.Windows.Forms.Application.StartupPath)
            expanded = ReplaceVariable(expanded, "%DESKTOP%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop))
            expanded = ReplaceVariable(expanded, "%DOCUMENTS%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments))
            expanded = ReplaceVariable(expanded, "%MYDOCUMENTS%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments))
            expanded = ReplaceVariable(expanded, "%DOWNLOADS%", System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.UserProfile), "Downloads"))
            expanded = ReplaceVariable(expanded, "%PROGRAMFILES%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles))
            expanded = ReplaceVariable(expanded, "%COMMONDOCUMENTS%", System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonDocuments))
            ' The Red Ink helper preserves an initial UNC path; only normal paths
            ' have repeated backslashes collapsed.
            If Not expanded.StartsWith("\\", System.StringComparison.Ordinal) Then
                expanded = System.Text.RegularExpressions.Regex.Replace(expanded, "\\{2,}", "\")
            End If
            If expanded.Length = 0 Then Return ""
            If System.IO.Path.IsPathRooted(expanded) Then Return System.IO.Path.GetFullPath(expanded).Trim()
            Return expanded.Trim()
        Catch ex As System.Exception
            AppLog.Write("path.expand.failed", ex)
            Return ""
        End Try
    End Function

    Private Shared Function ReplaceVariable(raw As String, variableName As String, replacement As String) As String
        If replacement Is Nothing Then
            If System.Text.RegularExpressions.Regex.IsMatch(raw, System.Text.RegularExpressions.Regex.Escape(variableName), System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                Throw New System.InvalidOperationException("Unavailable environment variable " & variableName)
            End If
            Return raw
        End If
        Return System.Text.RegularExpressions.Regex.Replace(raw, System.Text.RegularExpressions.Regex.Escape(variableName), replacement, System.Text.RegularExpressions.RegexOptions.IgnoreCase)
    End Function
End Class
