' Source of the resolution contract:
' SharedMethods.Settings.vb GetActiveConfigSource / ResolveRegistryConfigurationSource,
' SharedMethods.Constants.vb DefaultINIPaths / RegPath_IniPrio=False,
' SharedMethods.LoadConfig.vb redink.ini parser (16.10.2026).
Friend NotInheritable Class LibraryConfiguration
    Public Property Source As String = ""
    Public Property CentralPath As String = ""
    Public Property LocalPath As String = ""
    Public ReadOnly Property HasLocal As Boolean
        Get
            Return Not System.String.IsNullOrWhiteSpace(LocalPath)
        End Get
    End Property
End Class

Friend NotInheritable Class WordConfigReader
    Private Sub New()
    End Sub

    Public Shared Function ResolveSource() As String
        ' Equivalent to GetActiveConfigSource(context.RDV="Word") with RegPath_IniPrio=False.
        Dim wordDefault As String = RedInkPathExpander.Expand("%AppData%\Microsoft\Word\redink.ini")
        Dim regDirectory As String = ""
        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey("Software\Red Ink", False)
                If key IsNot Nothing Then regDirectory = System.Convert.ToString(key.GetValue("IniPath"), System.Globalization.CultureInfo.InvariantCulture)
            End Using
        Catch ex As System.Exception
            AppLog.Write("config.registry.read.failed", ex)
        End Try
        If System.IO.File.Exists(wordDefault) Then Return wordDefault
        If Not System.String.IsNullOrWhiteSpace(regDirectory) Then
            Dim raw As String = regDirectory.Trim().Replace(vbCr, "")
            If raw.Contains("://", System.StringComparison.Ordinal) Then Return raw
            ' Registry IniPath is a DIRECTORY, not an ini filename, matching Red Ink.
            Return System.IO.Path.Combine(RedInkPathExpander.Expand(raw), "redink.ini")
        End If
        Return wordDefault
    End Function

    Public Shared Async Function LoadAsync() As System.Threading.Tasks.Task(Of LibraryConfiguration)
        Dim source As String = ResolveSource()
        Dim content As String
        If source.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then
            Using handler As New System.Net.Http.HttpClientHandler With {.AllowAutoRedirect = False, .UseDefaultCredentials = True, .ClientCertificateOptions = System.Net.Http.ClientCertificateOption.Automatic}
                Using client As New System.Net.Http.HttpClient(handler) With {.Timeout = System.TimeSpan.FromSeconds(30)}
                    Using response As System.Net.Http.HttpResponseMessage = Await client.GetAsync(source, System.Net.Http.HttpCompletionOption.ResponseHeadersRead)
                        response.EnsureSuccessStatusCode()
                        content = Await response.Content.ReadAsStringAsync()
                    End Using
                End Using
            End Using
        ElseIf source.Contains("://", System.StringComparison.Ordinal) Then
            Throw New System.NotSupportedException("Nur Dateipfade und HTTPS-Konfigurationen werden unterstützt.")
        Else
            content = Await System.IO.File.ReadAllTextAsync(RedInkPathExpander.Expand(source))
        End If

        Dim dict As New System.Collections.Generic.Dictionary(Of String, String)(System.StringComparer.OrdinalIgnoreCase)
        For Each line As String In content.Split({vbCrLf, vbCr, vbLf}, System.StringSplitOptions.RemoveEmptyEntries)
            Dim trimmed As String = line.Trim()
            If trimmed.Length = 0 OrElse trimmed.StartsWith(";", System.StringComparison.Ordinal) Then Continue For
            Dim pair As String() = trimmed.Split(New Char() {"="c}, 2)
            If pair.Length = 2 Then dict(pair(0).Trim()) = pair(1).Trim()
        Next
        Dim central As String = ""
        Dim local As String = ""
        dict.TryGetValue("PromptLib", central)
        dict.TryGetValue("PromptLibLocal", local)
        Return New LibraryConfiguration With {
            .Source = source,
            .CentralPath = RedInkPathExpander.Expand(central),
            .LocalPath = RedInkPathExpander.Expand(local)
        }
    End Function
End Class
