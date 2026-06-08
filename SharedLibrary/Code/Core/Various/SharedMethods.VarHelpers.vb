' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedMethods.VarHelpers.vb
' Purpose: Provides small shared helper functions for common tasks used across the
'          library, including default INI path selection, file backup rename,
'          registry read/write, whitespace/newline normalization, environment
'          variable expansion for file paths, simple Base64/XOR encoding helpers,
'          and domain/workgroup checks.
'
' Architecture:
'  - INI path resolution: Selects a default INI path from `DefaultINIPaths` based
'    on substring matching of the provided key, then expands environment variables.
'  - File backup: Renames an existing file to the same path with a `.bak` suffix,
'    overwriting an existing `.bak` by deleting it first.
'  - Registry helpers: Writes and reads string values under a registry path,
'    selecting the hive by the path prefix.
'  - Text normalization: Removes CR/LF characters via `RemoveCR`.
'  - Path normalization: Expands environment variables and some additional
'    placeholder tokens, normalizes backslashes, and returns a full path.
'  - Encoding helpers: Decodes Base64 with URL-safe normalization; provides a
'    simple XOR encode/decode helper using a caller-provided term.
'  - Domain checks: Uses WMI (`Win32_ComputerSystem`) to read `Domain` and checks
'    it against the configured `alloweddomains` list.
'
' External Dependencies:
'  - Microsoft.Win32.Registry (registry access)
'  - System.Management (WMI query for domain/workgroup)
'  - System.Windows.Forms.MessageBox (error display)
'  - Internal helpers used here: `ShowCustomMessageBox`, `DefaultINIPaths`, `AN`,
'    and `alloweddomains`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.IO
Imports System.Management
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports SharedLibrary.SharedLibrary.SharedContext


Namespace SharedLibrary
    Partial Public Class SharedMethods


        ' Standard (dark theme):
        'Me.Icon = Icon.FromHandle(SharedMethods.GetLogoBitmap(_context, SharedMethods.LogoType.Standard).GetHicon())

        ' Medium (light theme):
        'Menu1.Image = SharedMethods.GetLogoBitmap(ThisAddIn._context, SharedMethods.LogoType.Medium)

        ' Large:
        'picLogo.Image = SharedMethods.GetLogoBitmap(_context, SharedMethods.LogoType.Large)

        ' SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard)


        ''' <summary>
        ''' Logo types that can be requested.
        ''' </summary>
        Public Enum LogoType
            Standard    ' Red_Ink_Logo (dark theme / default)
            Medium      ' Red_Ink_Logo_Medium (light theme)
            Large       ' Red_Ink_Logo_Large
        End Enum

        ''' <summary>
        ''' Cached logo path for Standard variant. Set during LoadConfig.
        ''' Supports file path, UNC path, or URL.
        ''' </summary>
        Public Shared INI_LogoPath_Cached As String = ""

        ''' <summary>
        ''' Cached logo path for Medium variant (light theme). Set during LoadConfig.
        ''' Supports file path, UNC path, or URL.
        ''' </summary>
        Public Shared INI_LogoPathMedium_Cached As String = ""

        ''' <summary>
        ''' Cached logo path for Large variant. Set during LoadConfig.
        ''' Supports file path, UNC path, or URL.
        ''' </summary>
        Public Shared INI_LogoPathLarge_Cached As String = ""

        ''' <summary>
        ''' Gets a logo bitmap from an external path (file, UNC, or URL) if configured and valid,
        ''' otherwise returns the appropriate embedded resource from SharedLibrary.
        ''' Uses the globally cached logo path values.
        ''' </summary>
        ''' <param name="logoType">Which logo type to retrieve.</param>
        ''' <returns>Valid Bitmap from external source or the embedded resource fallback.</returns>
        Public Shared Function GetLogoBitmap(logoType As LogoType, Optional RedInkLogo As Boolean = False) As System.Drawing.Bitmap

            If Not RedInkLogo Then
                Try

                    Dim logoPath As String = Nothing
                    Select Case logoType
                        Case LogoType.Standard
                            logoPath = ExpandEnvironmentVariables(If(INI_LogoPath_Cached, ""))
                        Case LogoType.Medium
                            logoPath = ExpandEnvironmentVariables(If(INI_LogoPathMedium_Cached, ""))
                        Case LogoType.Large
                            logoPath = ExpandEnvironmentVariables(If(INI_LogoPathLarge_Cached, ""))
                    End Select

                    If Not String.IsNullOrWhiteSpace(logoPath) Then
                        Dim bmp = LoadBitmapFromPath(logoPath)
                        If bmp IsNot Nothing AndAlso bmp.Width > 0 AndAlso bmp.Height > 0 Then
                            Return bmp
                        End If
                        bmp?.Dispose()
                    End If
                Catch
                End Try
            End If

            ' Fallback to embedded SharedLibrary resources
            Select Case logoType
                Case LogoType.Medium
                    Return My.Resources.Red_Ink_Logo_Medium
                Case LogoType.Large
                    Return My.Resources.Red_Ink_Logo_Large
                Case Else
                    Return My.Resources.Red_Ink_Logo
            End Select
        End Function

        Private Shared Function LoadBitmapFromPath(path As String) As System.Drawing.Bitmap
            Try
                ' URL (http/https)
                If path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) OrElse
           path.StartsWith("https://", StringComparison.OrdinalIgnoreCase) Then
                    Using client As New System.Net.WebClient()
                        Dim bytes = client.DownloadData(path)
                        Using ms As New System.IO.MemoryStream(bytes)
                            Return New System.Drawing.Bitmap(ms)
                        End Using
                    End Using
                End If

                ' UNC or local file
                If path.StartsWith("\\", StringComparison.Ordinal) OrElse System.IO.File.Exists(path) Then
                    Dim bytes = System.IO.File.ReadAllBytes(path)
                    Using ms As New System.IO.MemoryStream(bytes)
                        Return New System.Drawing.Bitmap(ms)
                    End Using
                End If
            Catch
            End Try
            Return Nothing
        End Function

        ''' <summary>
        ''' Returns the default INI path for a given key by selecting the first matching entry from <c>DefaultINIPaths</c>
        ''' and expanding environment variables in the resulting path.
        ''' </summary>
        ''' <param name="key">A key string used to select the default INI path (substring match).</param>
        ''' <returns>The expanded default INI path.</returns>
        Public Shared Function GetDefaultINIPath(ByVal key As String) As String

            For Each entry In DefaultINIPaths
                If key.Contains(entry.Key) Then
                    Return ExpandEnvironmentVariables(entry.Value)
                End If
            Next
            Return ExpandEnvironmentVariables(DefaultINIPaths.Values.First())
        End Function


        ''' <summary>
        ''' Renames the file at <paramref name="filePath"/> to <c>filePath + ".bak"</c>, deleting an existing
        ''' <c>.bak</c> file first if present.
        ''' </summary>
        ''' <param name="filePath">The path of the file to rename.</param>
        ''' <returns><c>True</c> on success; otherwise <c>False</c>.</returns>
        Public Shared Function RenameFileToBak(filePath As String) As Boolean
            Try
                ' Rename the file to a .bak file
                Dim bakFilePath As String = filePath & ".bak"
                If File.Exists(bakFilePath) Then
                    File.Delete(bakFilePath)
                End If
                File.Move(filePath, bakFilePath)
                Return True
            Catch ex As Exception
                ShowCustomMessageBox($"Error renaming file to .bak: {ex.Message}")
                Return False
            End Try

        End Function


        Public Shared Function GenerateRandomText(ByVal targetLength As Integer) As String
            If targetLength <= 0 Then
                Return String.Empty
            End If

            Dim words As String() = {
        "the", "and", "to", "of", "a", "in", "that", "is", "for", "on",
        "with", "as", "by", "this", "from", "or", "be", "are", "it", "an",
        "which", "at", "we", "can", "has", "have", "will", "not", "one",
        "all", "their", "more", "about", "when", "there", "use", "used",
        "such", "other", "some", "time", "each", "many", "these", "may",
        "like", "well", "very", "into", "over", "after", "before"
    }

            Dim random As New System.Random()
            Dim builder As New System.Text.StringBuilder(targetLength + 100)

            Dim currentParagraphLength As Integer = 0
            Dim paragraphTargetLength As Integer = random.Next(250, 500)

            While builder.Length < targetLength
                Dim word As String = words(random.Next(words.Length))

                If builder.Length > 0 Then
                    builder.Append(" ")
                    currentParagraphLength += 1
                End If

                builder.Append(word)
                currentParagraphLength += word.Length

                ' Insert paragraph breaks when paragraph size is reached
                If currentParagraphLength >= paragraphTargetLength AndAlso builder.Length < targetLength - 2 Then
                    builder.Append(System.Environment.NewLine)
                    builder.Append(System.Environment.NewLine)

                    currentParagraphLength = 0
                    paragraphTargetLength = random.Next(250, 500)
                End If
            End While

            ' Trim to exact requested length
            If builder.Length > targetLength Then
                builder.Length = targetLength
            End If

            Return builder.ToString()
        End Function



        ''' <summary>
        ''' Writes a string value to the Windows registry at <paramref name="regPath"/> using the default (unnamed) value.
        ''' </summary>
        ''' <param name="regPath">Registry path including hive name (e.g., <c>HKEY_CURRENT_USER\...</c>).</param>
        ''' <param name="regValue">Value to write; CR/LF is removed via <see cref="RemoveCR"/>.</param>
        Public Shared Sub WriteToRegistry(ByVal regPath As String, ByVal regValue As String)
            Try
                ' Remove carriage returns from the value
                regValue = RemoveCR(regValue)

                ' Split the registry path into hive and subkey
                Dim hiveName As String = regPath.Split("\"c)(0)
                Dim subKeyPath As String = String.Join("\", regPath.Split("\"c).Skip(1))

                Dim registryHive As RegistryKey

                ' Determine the appropriate registry hive
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        registryHive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        registryHive = Registry.LocalMachine
                    Case "HKEY_CLASSES_ROOT"
                        registryHive = Registry.ClassesRoot
                    Case "HKEY_USERS"
                        registryHive = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        registryHive = Registry.CurrentConfig
                    Case Else
                        Throw New ArgumentException("Unsupported registry hive: " & hiveName)
                End Select

                ' Write the value to the registry
                Using subKey As RegistryKey = registryHive.CreateSubKey(subKeyPath, True)
                    If subKey Is Nothing Then
                        Throw New Exception("Unable to open or create the registry key at: " & regPath)
                    End If
                    subKey.SetValue("", regValue, RegistryValueKind.String)
                End Using

                ShowCustomMessageBox($"Written value '{regValue}' to the registry at '{regPath}.'")

            Catch ex As Exception
                MessageBox.Show($"Error: Unable to write to the registry at '{regPath}'. {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub


        ''' <summary>
        ''' Reads a value from the Windows registry.
        ''' </summary>
        ''' <param name="registryPath">Registry path including hive name (e.g., <c>HKEY_CURRENT_USER\...</c>).</param>
        ''' <param name="valueName">Name of the value to read.</param>
        ''' <param name="suppressErrors">If <c>True</c>, suppresses error message boxes and returns an empty string on failure.</param>
        ''' <returns>The registry value as a string with CR/LF removed, or an empty string on failure/not found.</returns>
        Public Shared Function GetFromRegistry(registryPath As String, valueName As String, Optional suppressErrors As Boolean = False) As String
            Try
                ' Split the registry path into hive and subkey
                Dim hiveName As String = registryPath.Split("\"c)(0)
                Dim subKeyPath As String = registryPath.Substring(hiveName.Length + 1)

                ' Determine the registry hive
                Dim hive As RegistryKey = Nothing
                Select Case hiveName.ToUpper()
                    Case "HKEY_CURRENT_USER"
                        hive = Registry.CurrentUser
                    Case "HKEY_LOCAL_MACHINE"
                        hive = Registry.LocalMachine
                    Case "HKEY_CLASSES_ROOT"
                        hive = Registry.ClassesRoot
                    Case "HKEY_USERS"
                        hive = Registry.Users
                    Case "HKEY_CURRENT_CONFIG"
                        hive = Registry.CurrentConfig
                    Case Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - invalid registry hive: " & hiveName, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                End Select

                ' Open the subkey and retrieve the value
                Using subKey As RegistryKey = hive.OpenSubKey(subKeyPath)
                    If subKey IsNot Nothing Then
                        Return RemoveCR(subKey.GetValue(valueName, Nothing)?.ToString())
                    Else
                        If Not suppressErrors Then
                            MessageBox.Show("Error in GetFromRegistry - Registry key not found: " & subKeyPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        End If
                        Return ""
                    End If
                End Using

            Catch ex As System.Exception
                If Not suppressErrors Then
                    MessageBox.Show("An error occurred: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Removes CR/LF characters from a string and trims leading/trailing whitespace.
        ''' </summary>
        ''' <param name="inputtext">Input text.</param>
        ''' <returns>Text with CR/LF removed, or an empty string if <paramref name="inputtext"/> is <c>Nothing</c>.</returns>
        Public Shared Function RemoveCR(ByVal inputtext As String) As String
            If inputtext IsNot Nothing Then
                inputtext = inputtext.Trim()
                inputtext = inputtext.Replace(vbCr, "")
                inputtext = inputtext.Replace(vbLf, "")
                inputtext = inputtext.Replace(vbCrLf, "")
                inputtext = inputtext.Trim()
            Else
                inputtext = ""
            End If
            Return inputtext
        End Function

        ''' <summary>
        ''' Returns <c>True</c> if <paramref name="str"/> is <c>Nothing</c>, empty, or consists only of whitespace.
        ''' </summary>
        ''' <param name="str">Input string.</param>
        ''' <returns><c>True</c> if the string is <c>Nothing</c>, empty, or whitespace; otherwise <c>False</c>.</returns>
        Public Shared Function IsEmptyOrBlank(ByVal str As String) As Boolean
            ' Check if the string is empty or consists only of whitespace
            Return String.IsNullOrWhiteSpace(str)
        End Function



        Public Shared Function GetDictionaryFilePath(ByVal context As ISharedContext,
                                                     Optional ByVal localDictionary As Boolean = False) As String
            If context Is Nothing Then Return ""

            Dim configuredPath As String = If(localDictionary, context.INI_DictionaryPathLocal, context.INI_DictionaryPath)
            Return ExpandEnvironmentVariables(configuredPath)
        End Function

        Public Shared Function GetTranslationDictionaryText(ByVal context As ISharedContext) As String
            If context Is Nothing Then Return ""

            Dim globalDictionary As String = ReadTranslationDictionaryFile(GetDictionaryFilePath(context, False))
            Dim userDictionary As String = ReadTranslationDictionaryFile(GetDictionaryFilePath(context, True))

            Dim parts As New System.Collections.Generic.List(Of String)

            If Not String.IsNullOrWhiteSpace(globalDictionary) Then
                parts.Add("<globaldictionary>" & globalDictionary & "</globaldictionary>")
            End If

            If Not String.IsNullOrWhiteSpace(userDictionary) Then
                Dim userPrefix As String = If(parts.Count > 0, "(the userdictionary takes precedence) ", "")
                parts.Add("<userdictionary>" & userPrefix & userDictionary & "</userdictionary>")
            End If

            Return String.Join(" ", parts)
        End Function

        Public Shared Function PromptForTargetLanguage(ByVal prompt As String,
                                                       ByVal header As String,
                                                       ByVal defaultLanguage As String,
                                                       ByVal context As ISharedContext) As String
            Dim inputParams() As InputParameter = {
                New InputParameter("", If(defaultLanguage, ""))
            }

            Dim confirmed As Boolean = ShowCustomVariableInputForm(
                prompt,
                header,
                inputParams,
                "Edit User Dictionary",
                Sub()
                    EditUserDictionaryFile(context)
                End Sub,
                False)

            If Not confirmed Then Return ""

            Return If(System.Convert.ToString(inputParams(0).Value), "").Trim()
        End Function

        Public Shared Sub EditUserDictionaryFile(ByVal context As ISharedContext)
            Dim filePath As String = EnsureUserDictionaryFileExists(context)
            If String.IsNullOrWhiteSpace(filePath) Then Return

            ShowTextFileEditor(filePath, "Edit user dictionary: " & filePath, False, context)
        End Sub

        Private Shared Function EnsureUserDictionaryFileExists(ByVal context As ISharedContext) As String
            Dim filePath As String = GetDictionaryFilePath(context, True)

            If String.IsNullOrWhiteSpace(filePath) Then
                ShowCustomMessageBox("No user dictionary file is configured.")
                Return ""
            End If

            Try
                filePath = Path.GetFullPath(filePath)

                Dim directoryPath As String = Path.GetDirectoryName(filePath)
                If String.IsNullOrWhiteSpace(directoryPath) Then
                    ShowCustomMessageBox("Invalid user dictionary file path:" & vbCrLf & filePath)
                    Return ""
                End If

                If Not Directory.Exists(directoryPath) Then
                    Directory.CreateDirectory(directoryPath)
                End If

                If Not File.Exists(filePath) Then
                    Dim initialContent As String =
                        "; This is the user dictionary." & vbCrLf &
                        "; It can contain instructions on how to translate certain words or expressions." & vbCrLf &
                        "; Example: Translate ""share purchase agreement"" as ""Aktienkaufvertrag""." & vbCrLf

                    File.WriteAllText(filePath, initialContent, New System.Text.UTF8Encoding(True))
                End If

                Return filePath
            Catch ex As Exception
                ShowCustomMessageBox("Could not create the user dictionary file:" & vbCrLf & ex.Message)
                Return ""
            End Try
        End Function

        Private Shared Function ReadTranslationDictionaryFile(ByVal filePath As String) As String
            If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then Return ""

            Dim content As String = ReadTextFile(filePath, False)
            If String.IsNullOrWhiteSpace(content) Then Return ""

            Dim normalized As String = content.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)
            Dim sb As New System.Text.StringBuilder()

            For Each line As String In normalized.Split(New String() {vbLf}, StringSplitOptions.None)
                If line.TrimStart().StartsWith(";", StringComparison.Ordinal) Then
                    Continue For
                End If

                If sb.Length > 0 Then
                    sb.Append(vbCrLf)
                End If

                sb.Append(line)
            Next

            Return sb.ToString().Trim()
        End Function


        ''' <summary>
        ''' Expands environment variables and selected placeholder tokens in <paramref name="filePath"/>,
        ''' normalizes the resulting path, and returns its full path.
        ''' </summary>
        ''' <param name="filePath">Path that may contain environment variables (e.g., <c>%APPDATA%</c>).</param>
        ''' <returns>The expanded full path, or an empty string on failure.</returns>
        Public Shared Function ExpandEnvironmentVariables(ByVal filePath As String) As String
            ' Handle null/empty input early
            If String.IsNullOrWhiteSpace(filePath) Then Return ""

            ' Trim whitespace from input before processing
            filePath = filePath.Trim()

            ' Start with the input path
            Dim expandedPath As String = Environment.ExpandEnvironmentVariables(filePath)

            Try

                ' Remove any preceding and trailing quotation marks
                expandedPath = expandedPath.Trim(""""c)

                ' Expand known variables using Environment.GetEnvironmentVariable and ensure proper path format
                expandedPath = Regex.Replace(expandedPath, "%APPDATA%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%USERPROFILE%", Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%WINDIR%", Path.Combine(Environment.GetEnvironmentVariable("WINDIR")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%TEMP%", Path.Combine(Path.GetTempPath()), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%HOMEPATH%", Path.Combine(Environment.GetEnvironmentVariable("HOMEPATH")), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%APPSTARTUPPATH%", Path.Combine(System.Windows.Forms.Application.StartupPath), RegexOptions.IgnoreCase)
                expandedPath = Regex.Replace(expandedPath, "%DESKTOP%", Environment.GetFolderPath(Environment.SpecialFolder.Desktop), RegexOptions.IgnoreCase)

                ' Clean up any potential double backslashes (but preserve UNC paths)
                If Not expandedPath.StartsWith("\\") Then
                    expandedPath = Regex.Replace(expandedPath, "\\{2,}", "\")
                End If

                ' Only normalize with GetFullPath if the path is already absolute.
                ' Relative paths should remain relative - the caller knows the correct base directory.
                If String.IsNullOrEmpty(expandedPath) Then
                    Return ""
                ElseIf Path.IsPathRooted(expandedPath) Then
                    Return Path.GetFullPath(expandedPath).Trim()
                Else
                    Return expandedPath.Trim()
                End If

            Catch ex As System.Exception
                ' Return empty string on failure
                Return ""
            End Try

        End Function


        ''' <summary>
        ''' Decodes a Base64 string to bytes, normalizing whitespace and URL-safe Base64 variants.
        ''' </summary>
        ''' <param name="base64String">Input Base64 (may contain whitespace/newlines and may be URL-safe).</param>
        ''' <returns>Decoded bytes, or <c>Nothing</c> on failure.</returns>
        Public Shared Function DecodeBase64(ByVal base64String As String) As Byte()
            Try
                ' Normalize the input: remove whitespaces and line breaks
                base64String = base64String.Replace(vbCrLf, "").Replace(vbLf, "").Replace(vbCr, "").Replace(" ", "")

                ' Convert URL-safe Base64 to standard Base64 if input is URL-safe
                base64String = base64String.Replace("-", "+").Replace("_", "/")

                ' Add padding
                While (base64String.Length Mod 4) <> 0
                    base64String &= "="
                End While

                ' Decode the Base64 string
                Return System.Convert.FromBase64String(base64String)
            Catch ex As System.Exception
                ' Return Nothing on failure
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Decodes an XOR-encoded and Base64-encoded string using <paramref name="pTerm"/> as the XOR key material.
        ''' </summary>
        ''' <param name="encodedText">The Base64 input text (whitespace/newlines are removed before decoding).</param>
        ''' <param name="pTerm">XOR key term used for decoding.</param>
        ''' <returns>Decoded text, or the literal string <c>"Error: Invalid Base64 input"</c> if Base64 decoding fails.</returns>
        Public Shared Function DecodeString(ByVal encodedText As String, ByVal pTerm As String) As String
            encodedText = If(encodedText, "")
            pTerm = If(pTerm, "")

            If String.IsNullOrEmpty(pTerm) Then
                Return "Error: Missing CodeBasis"
            End If

            encodedText = encodedText.Replace("\n", "")
            encodedText = encodedText.Replace(vbCr, "").Replace(vbLf, "")
            encodedText = encodedText.Replace(" ", "")

            Dim encryptedBytes As Byte() = DecodeBase64(encodedText)
            If encryptedBytes Is Nothing Then
                Return "Error: Invalid Base64 input"
            End If

            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            If pTermBytes.Length = 0 Then
                Return "Error: Missing CodeBasis"
            End If

            Dim decryptedBytes(encryptedBytes.Length - 1) As Byte

            For i As Integer = 0 To encryptedBytes.Length - 1
                decryptedBytes(i) = encryptedBytes(i) Xor pTermBytes(i Mod pTermBytes.Length)
            Next

            Try
                Return System.Text.Encoding.UTF8.GetString(decryptedBytes)
            Catch
                Return System.Text.Encoding.ASCII.GetString(decryptedBytes)
            End Try
        End Function



        ''' <summary>
        ''' XOR-encodes <paramref name="inputText"/> using <paramref name="pTerm"/> and returns the result as Base64.
        ''' </summary>
        ''' <param name="inputText">Plain text to encode.</param>
        ''' <param name="pTerm">XOR key term used for encoding.</param>
        ''' <returns>Base64-encoded XOR output.</returns>
        Public Shared Function CodeString(ByVal inputText As String, ByVal pTerm As String) As String
            Dim inputBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(inputText)
            Dim pTermBytes() As Byte = System.Text.Encoding.UTF8.GetBytes(pTerm)
            Dim encryptedBytes(inputBytes.Length - 1) As Byte

            Dim inputLength As Integer = inputBytes.Length
            Dim pTermLength As Integer = pTermBytes.Length

            ' Encrypt each byte with XOR operation
            For i As Integer = 0 To inputBytes.Length - 1
                encryptedBytes(i) = inputBytes(i) Xor pTermBytes(i Mod pTermLength)
            Next

            ' Convert encrypted bytes to Base64
            Return System.Convert.ToBase64String(encryptedBytes)
        End Function

        ''' <summary>
        ''' Retrieves the computer's domain/workgroup name via WMI (<c>Win32_ComputerSystem.Domain</c>).
        ''' </summary>
        ''' <returns>The domain/workgroup name, or an empty string on failure.</returns>
        Public Shared Function GetDomain() As String
            Try
                ' Initialize a WMI query to get the Domain property from Win32_ComputerSystem
                Dim searcher As New ManagementObjectSearcher("SELECT Domain FROM Win32_ComputerSystem")
                Dim strDomain As String = String.Empty

                ' Execute the query and retrieve the result
                For Each queryObj As ManagementObject In searcher.Get()
                    If queryObj("Domain") IsNot Nothing Then
                        strDomain = queryObj("Domain").ToString()
                    End If
                Next

                ' If the domain is not retrieved, return an appropriate message
                If String.IsNullOrEmpty(strDomain) Then
                    MessageBox.Show($"Error in GetDomain - unable to determine the domain name or workgroup.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    strDomain = ""
                End If

                Return strDomain
            Catch ex As System.Exception
                MessageBox.Show($"Error in GetDomain - Error retrieving domain or workgroup: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return String.Empty
            End Try
        End Function

        ''' <summary>
        ''' Returns <c>True</c> when <c>alloweddomains</c> is configured and the current domain/workgroup name
        ''' (from <see cref="GetDomain"/>) is not contained in that list.
        ''' </summary>
        ''' <returns><c>True</c> if the current domain/workgroup is not allowed; otherwise <c>False</c>.</returns>
        Public Shared Function WrongDomain() As Boolean

            Dim strDomain As String = GetDomain() ' Current domain of the computer
            Dim domainList() As String
            Dim domainFound As Boolean = False

            If Not String.IsNullOrEmpty(alloweddomains) Then
                ' Convert the list of allowed domains into an array
                domainList = alloweddomains.Split(","c)

                ' Check if the current domain is in the allowed list
                For Each domain In domainList
                    If strDomain.Equals(domain.Trim(), StringComparison.OrdinalIgnoreCase) Then
                        domainFound = True
                        Exit For
                    End If
                Next

                ' If the domain is not in the list of allowed domains
                If Not domainFound Then
                    ShowCustomMessageBox($"This copy of {AN} may not be executed in this network environment (which is '{strDomain}'). The domain has to be added to the code by your administrator.")
                    Return True
                Else
                    Return False
                End If
            Else
                Return False
            End If
        End Function


        ''' <summary>
        ''' Uppercases the first character of a string (culture-invariant) and leaves the rest unchanged.
        ''' </summary>
        Public Shared Function StartWithUpcase(value As String) As String
            If String.IsNullOrEmpty(value) Then Return If(value, "")

            Dim first As Char = value(0)
            Dim upperFirst As Char = Char.ToUpperInvariant(first)

            If value.Length = 1 Then
                Return upperFirst.ToString()
            End If

            If upperFirst = first Then
                Return value
            End If

            Return upperFirst & value.Substring(1)
        End Function

        ''' <summary>
        ''' Checks whether an ImageGeneration special task model is configured and available.
        ''' Probes the alternate model file without permanently switching the active model.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <param name="hasObjectCall">
        ''' Output: <c>True</c> if the ImageGeneration model also has an APICall_Object configured,
        ''' indicating it supports file/image input (e.g. for image editing).
        ''' </param>
        ''' <returns><c>True</c> if an ImageGeneration model is available; otherwise <c>False</c>.</returns>
        Public Shared Function IsImageGenerationAvailable(context As SharedContext.ISharedContext,
                                                          Optional ByRef hasObjectCall As Boolean = False) As Boolean
            hasObjectCall = False
            If context Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then Return False

            Dim scope = CaptureModelConfigScope(context)
            Dim available As Boolean = False

            Try
                available = GetSpecialTaskModel(context, context.INI_AlternateModelPath, "ImageGeneration")
                If available Then
                    hasObjectCall = Not String.IsNullOrWhiteSpace(context.INI_APICall_Object_2)
                End If
            Catch
            Finally
                RestoreModelConfigScope(context, scope)
            End Try

            Return available
        End Function


        ''' <summary>
        ''' Checks whether a WebGrounding special task model is configured and available.
        ''' Probes the alternate model file without permanently switching the active model.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns><c>True</c> if a WebGrounding model is available; otherwise <c>False</c>.</returns>
        Public Shared Function IsWebGroundingAvailable(context As SharedContext.ISharedContext) As Boolean
            If context Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then Return False

            Dim scope = CaptureModelConfigScope(context)
            Dim available As Boolean = False

            Try
                available = GetSpecialTaskModel(context, context.INI_AlternateModelPath, "WebGrounding")
            Catch
            Finally
                RestoreModelConfigScope(context, scope)
            End Try

            Return available
        End Function

        ''' <summary>
        ''' Checks whether a DeepResearch special task model is configured and available.
        ''' Probes the alternate model file without permanently switching the active model.
        ''' </summary>
        ''' <param name="context">Shared context containing model and API configuration.</param>
        ''' <returns><c>True</c> if a DeepResearch model is available; otherwise <c>False</c>.</returns>
        Public Shared Function IsDeepResearchAvailable(context As SharedContext.ISharedContext) As Boolean
            If context Is Nothing Then Return False
            If String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then Return False

            Dim scope = CaptureModelConfigScope(context)
            Dim available As Boolean = False

            Try
                available = GetSpecialTaskModel(context, context.INI_AlternateModelPath, "DeepResearch")
            Catch
            Finally
                RestoreModelConfigScope(context, scope)
            End Try

            Return available
        End Function

        ''' <summary>
        ''' Ensures a form is fully reachable on an active screen.
        ''' If monitors/resolution changed, the form is resized as needed and clamped into the
        ''' nearest working area so that it remains accessible.
        ''' </summary>
        ''' <param name="frm">The form to check and reposition if necessary.</param>
        Public Shared Sub EnsureVisibleOnScreen(frm As System.Windows.Forms.Form)
            If frm Is Nothing OrElse frm.IsDisposed Then Return

            Dim minWidth As Integer = 0
            Dim minHeight As Integer = 0

            If frm.MinimumSize <> System.Drawing.Size.Empty Then
                minWidth = Math.Max(0, frm.MinimumSize.Width)
                minHeight = Math.Max(0, frm.MinimumSize.Height)
            End If

            Dim newWidth As Integer = frm.Width
            Dim newHeight As Integer = frm.Height

            If newWidth <= 0 Then newWidth = Math.Max(minWidth, 200)
            If newHeight <= 0 Then newHeight = Math.Max(minHeight, 120)

            If minWidth > 0 Then newWidth = Math.Max(newWidth, minWidth)
            If minHeight > 0 Then newHeight = Math.Max(newHeight, minHeight)

            Dim targetBounds As New System.Drawing.Rectangle(frm.Left, frm.Top, newWidth, newHeight)
            Dim area As System.Drawing.Rectangle = System.Windows.Forms.Screen.FromRectangle(targetBounds).WorkingArea

            If area.Width > 0 Then
                If minWidth > 0 AndAlso area.Width >= minWidth Then
                    newWidth = Math.Max(minWidth, Math.Min(newWidth, area.Width))
                Else
                    newWidth = Math.Min(newWidth, area.Width)
                End If
            End If

            If area.Height > 0 Then
                If minHeight > 0 AndAlso area.Height >= minHeight Then
                    newHeight = Math.Max(minHeight, Math.Min(newHeight, area.Height))
                Else
                    newHeight = Math.Min(newHeight, area.Height)
                End If
            End If

            Dim newLeft As Integer
            If newWidth >= area.Width Then
                newLeft = area.Left
            Else
                newLeft = Math.Min(Math.Max(frm.Left, area.Left), area.Right - newWidth)
            End If

            Dim newTop As Integer
            If newHeight >= area.Height Then
                newTop = area.Top
            Else
                newTop = Math.Min(Math.Max(frm.Top, area.Top), area.Bottom - newHeight)
            End If

            frm.SetBounds(newLeft, newTop, newWidth, newHeight)
        End Sub

    End Class

End Namespace

