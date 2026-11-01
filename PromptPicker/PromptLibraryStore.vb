' Mirrors SharedMethods.PromptLibrary.vb LoadPromptEntriesIntoList and AppendPromptLibraryEntry.
Friend NotInheritable Class PromptEntry
    Public Property Id As Integer
    Public Property Title As String = ""
    Public Property Prompt As String = ""
    Public Property Category As String = ""
    Public Property Source As String = ""
End Class

Friend NotInheritable Class PromptLibraryStore
    Private Sub New()
    End Sub

    Public Shared Function Load(config As LibraryConfiguration) As System.Collections.Generic.List(Of PromptEntry)
        Dim result As New System.Collections.Generic.List(Of PromptEntry)()
        ' Red Ink loads LOCAL before CENTRAL, with " (local)" appended to displayed titles.
        ReadFile(config.LocalPath, "local", " (local)", result)
        ReadFile(config.CentralPath, "central", "", result)
        For i As Integer = 0 To result.Count - 1
            result(i).Id = i
        Next
        Return result
    End Function

    Private Shared Sub ReadFile(path As String, source As String, suffix As String, output As System.Collections.Generic.List(Of PromptEntry))
        If System.String.IsNullOrWhiteSpace(path) Then Return
        Try
            If Not System.IO.File.Exists(path) Then Return
            Dim currentCategory As String = Nothing
            For Each line As String In System.IO.File.ReadAllLines(path)
                Dim value As String = line.Trim()
                If value.Length = 0 OrElse value.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
                If value.StartsWith("</", System.StringComparison.Ordinal) AndAlso value.EndsWith(">", System.StringComparison.Ordinal) Then
                    currentCategory = Nothing
                    Continue For
                End If
                If value.StartsWith("<", System.StringComparison.Ordinal) AndAlso value.EndsWith(">", System.StringComparison.Ordinal) AndAlso Not value.Contains("|", System.StringComparison.Ordinal) Then
                    Dim name As String = value.Substring(1, value.Length - 2).Trim()
                    If name.Length > 0 AndAlso Not name.StartsWith("/", System.StringComparison.Ordinal) Then
                        currentCategory = name
                        Continue For
                    End If
                End If
                Dim firstDelimiter As Integer = value.IndexOf("|"c)
                If firstDelimiter < 0 Then Continue For
                output.Add(New PromptEntry With {
                    .Title = value.Substring(0, firstDelimiter).Trim() & suffix,
                    .Prompt = value.Substring(firstDelimiter + 1).Trim(),
                    .Category = If(System.String.IsNullOrWhiteSpace(currentCategory), "(Uncategorized)", currentCategory.Trim()),
                    .Source = source
                })
            Next
        Catch ex As System.Exception
            ' Original picker ignores unreadable individual libraries and shows the rest.
            AppLog.Write("library.read.failed " & source, ex)
        End Try
    End Sub

    Public Shared ReadOnly Property CanSaveLocalPlaceholder As String
        Get
            Return "Nur bei konfiguriertem PromptLibLocal verfügbar"
        End Get
    End Property

    Public Shared Function CanSaveLocal(config As LibraryConfiguration) As Boolean
        If Not config.HasLocal Then Return False
        If config.LocalPath.Contains("://", System.StringComparison.Ordinal) Then Return False
        Try
            ' Never write into the central library even if both INI keys point to
            ' the same file (or to a direct file symlink alias).
            If Not System.String.IsNullOrWhiteSpace(config.CentralPath) AndAlso
               Not config.CentralPath.Contains("://", System.StringComparison.Ordinal) AndAlso
               System.String.Equals(CanonicalFilePath(config.LocalPath), CanonicalFilePath(config.CentralPath), System.StringComparison.OrdinalIgnoreCase) Then Return False
            Dim directory As String = System.IO.Path.GetDirectoryName(config.LocalPath)
            If System.String.IsNullOrWhiteSpace(directory) Then Return False
            If System.IO.File.Exists(config.LocalPath) Then
                Return (System.IO.File.GetAttributes(config.LocalPath) And System.IO.FileAttributes.ReadOnly) = CType(0, System.IO.FileAttributes)
            End If
            Return True  ' Configuration already explicitly defines the target; creation is permitted.
        Catch
            Return False
        End Try
    End Function

    Private Shared Function CanonicalFilePath(path As String) As String
        Dim full As String = System.IO.Path.GetFullPath(path)
        If System.IO.File.Exists(full) Then
            Dim target As System.IO.FileSystemInfo = New System.IO.FileInfo(full).ResolveLinkTarget(True)
            If target IsNot Nothing Then Return System.IO.Path.GetFullPath(target.FullName)
        End If
        Return full
    End Function

    Public Shared Sub SaveLocal(config As LibraryConfiguration, title As String, prompt As String, category As String)
        ' Strict boundary: absolutely no fallback to central; never writes redink.ini.
        If Not CanSaveLocal(config) Then Throw New System.InvalidOperationException("Keine beschreibbare lokale Prompt Library in PromptLibLocal konfiguriert.")
        title = Sanitize(title)
        category = Sanitize(category).Replace("<", "").Replace(">", "")
        prompt = If(prompt, "").Replace(vbCrLf, "\n").Replace(vbCr, "\n").Replace(vbLf, "\n").Replace("|", "¦").Trim()
        If title.Length = 0 OrElse prompt.Length = 0 Then Throw New System.InvalidOperationException("Titel und Prompt dürfen nicht leer sein.")
        Dim folder As String = System.IO.Path.GetDirectoryName(config.LocalPath)
        If System.String.IsNullOrWhiteSpace(folder) Then Throw New System.InvalidOperationException("Ungültiger lokaler Bibliothekspfad.")
        System.IO.Directory.CreateDirectory(folder)
        Dim existed As Boolean = System.IO.File.Exists(config.LocalPath)
        Dim sb As New System.Text.StringBuilder()
        If Not existed Then
            sb.AppendLine("; Local prompt library").AppendLine("; Format: Title|Prompt").AppendLine("; Optional categories: <Category Name> ... </Category Name>").AppendLine()
        End If
        If existed AndAlso New System.IO.FileInfo(config.LocalPath).Length > 0 Then sb.AppendLine()
        If category.Length > 0 Then sb.AppendLine("<" & category & ">")
        sb.AppendLine(title & "|" & prompt)
        If category.Length > 0 Then sb.AppendLine("</" & category & ">")
        System.IO.File.AppendAllText(config.LocalPath, sb.ToString(), New System.Text.UTF8Encoding(False))
    End Sub

    Private Shared Function Sanitize(text As String) As String
        Dim value As String = If(text, "").Replace(vbCr, " ").Replace(vbLf, " ")
        Return System.Text.RegularExpressions.Regex.Replace(value, "\s+", " ").Replace("|", "¦").Trim()
    End Function
End Class
