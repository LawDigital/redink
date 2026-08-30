' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Resolves the externally provisioned Playwright runtime used by BrowserTools.
' The runtime is intentionally independent from Word/Outlook ClickOnce deployment.

Option Explicit On
Option Strict On
Option Infer On

Namespace Agents

    Friend NotInheritable Class PlaywrightRuntimeResolution
        Public Property SourceRoot As System.String
        Public Property EffectiveRoot As System.String
        Public Property PlaywrightDirectory As System.String
        Public Property BrowsersDirectory As System.String
        Public Property RuntimeVersion As System.Version
        Public Property RuntimeVersionText As System.String
        Public Property UsesLocalCache As System.Boolean
    End Class

    Friend NotInheritable Class PlaywrightRuntimeResolver
        Private Const PlaywrightFolderName As System.String = ".playwright"
        Private Const LocalCacheProductFolder As System.String = "RedInk"
        Private Const LocalCachePlaywrightFolder As System.String = "Playwright"
        Private Const RuntimeManifestFileName As System.String = "redink-playwright-runtime.json"
        Private Const InvalidRuntimeMarkerFileName As System.String = ".redink-invalid"
        Private Shared ReadOnly BackgroundCacheLock As New System.Object()
        Private Shared ReadOnly BackgroundCacheTasks As New System.Collections.Generic.Dictionary(Of System.String, System.Threading.Tasks.Task)(System.StringComparer.OrdinalIgnoreCase)
        Private Shared ReadOnly InvalidLocalCacheLock As New System.Object()
        Private Shared ReadOnly InvalidLocalCacheRoots As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

        Private Sub New()
        End Sub

        Public Shared Function TryResolve(
            configuredPath As System.String,
            useLocalCache As System.Boolean,
            prepareLocalCache As System.Boolean,
            ByRef resolution As PlaywrightRuntimeResolution,
            ByRef errorMessage As System.String,
            Optional cachePreparationStarted As System.Action = Nothing
        ) As System.Boolean
            resolution = Nothing
            errorMessage = System.String.Empty

            Dim expandedPath As System.String = ExpandConfiguredPath(configuredPath)
            If System.String.IsNullOrWhiteSpace(expandedPath) Then
                errorMessage = "PlayWrightPath is not configured. Browser tools require an externally provisioned Playwright runtime."
                Return False
            End If

            Dim expectedVersion As System.Version = GetExpectedPlaywrightVersion()
            If expectedVersion Is Nothing Then
                errorMessage = "Could not determine the Microsoft.Playwright NuGet/driver release version. Generic 1.0.x assembly metadata is intentionally rejected; browser tools are unavailable rather than selecting an arbitrary runtime."
                Return False
            End If

            ' Prefer an already prepared compatible local cache. This path is fast and avoids
            ' the network share once the one-time background provisioning has completed.
            If useLocalCache Then
                Dim cached As PlaywrightRuntimeResolution = TryResolveExistingLocalCache(expectedVersion)
                If cached IsNot Nothing Then
                    resolution = cached
                    Return True
                End If
            End If

            Dim sourceCandidates As New System.Collections.Generic.List(Of PlaywrightRuntimeResolution)()
            Dim sourceAccessible As System.Boolean = System.IO.Directory.Exists(expandedPath)
            If sourceAccessible Then
                sourceCandidates = DiscoverCandidates(expandedPath)
            End If

            If Not sourceAccessible Then
                errorMessage = "Configured PlayWrightPath does not exist or is not accessible: " & expandedPath
                If useLocalCache Then errorMessage &= ". No compatible local Red Ink Playwright cache is available."
                Return False
            End If

            If sourceCandidates.Count = 0 Then
                errorMessage = "No valid Playwright runtime was found below PlayWrightPath. Expected a .playwright directory containing node\win32_x64\node.exe and package\cli.js."
                Return False
            End If

            Dim selected As PlaywrightRuntimeResolution = SelectCompatibleCandidate(sourceCandidates, expectedVersion)
            If selected Is Nothing Then
                Dim available As System.String = BuildAvailableVersionList(sourceCandidates)
                errorMessage = "No Playwright runtime compatible with Microsoft.Playwright " & FormatVersion(expectedVersion) & " was found. Available runtime versions: " & available & ". Newer runtimes may coexist, but Red Ink only selects a runtime with the same major/minor client version."
                Return False
            End If

            ' Never block tool discovery or the first browser call on a potentially large copy.
            ' The source runtime is immediately usable; when requested, provision a local copy
            ' in the background and prefer it automatically on later resolutions.
            If useLocalCache AndAlso prepareLocalCache Then
                StartLocalCachePreparation(selected, cachePreparationStarted)
            End If

            selected.UsesLocalCache = False
            resolution = selected
            Return True
        End Function

        Private Shared Function TryResolveExistingLocalCache(expectedVersion As System.Version) As PlaywrightRuntimeResolution
            Try
                Dim cacheBase As System.String = GetLocalCacheBasePath()
                If Not System.IO.Directory.Exists(cacheBase) Then Return Nothing
                Dim candidates As System.Collections.Generic.List(Of PlaywrightRuntimeResolution) = DiscoverCandidates(cacheBase)
                Dim selected As PlaywrightRuntimeResolution = SelectCompatibleCandidate(candidates, expectedVersion)
                If selected IsNot Nothing Then selected.UsesLocalCache = True
                Return selected
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return Nothing
            End Try
        End Function

        Private Shared Function ExpandConfiguredPath(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then Return System.String.Empty
            Try
                Return System.IO.Path.GetFullPath(System.Environment.ExpandEnvironmentVariables(value.Trim()))
            Catch ex As System.Exception
                Return System.Environment.ExpandEnvironmentVariables(value.Trim())
            End Try
        End Function

        Private Shared Function DiscoverCandidates(configuredPath As System.String) As System.Collections.Generic.List(Of PlaywrightRuntimeResolution)
            Dim result As New System.Collections.Generic.List(Of PlaywrightRuntimeResolution)()
            Dim seen As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

            AddCandidate(result, seen, configuredPath)

            Try
                For Each child As System.String In System.IO.Directory.GetDirectories(configuredPath)
                    AddCandidate(result, seen, child)
                Next
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Return result
        End Function

        Private Shared Sub AddCandidate(
            result As System.Collections.Generic.List(Of PlaywrightRuntimeResolution),
            seen As System.Collections.Generic.HashSet(Of System.String),
            candidatePath As System.String
        )
            If System.String.IsNullOrWhiteSpace(candidatePath) Then Return

            Dim root As System.String = candidatePath
            Dim playwrightDirectory As System.String

            If System.String.Equals(System.IO.Path.GetFileName(candidatePath.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)), PlaywrightFolderName, System.StringComparison.OrdinalIgnoreCase) Then
                playwrightDirectory = candidatePath
                Dim parent As System.IO.DirectoryInfo = System.IO.Directory.GetParent(candidatePath)
                If parent Is Nothing Then Return
                root = parent.FullName
            Else
                playwrightDirectory = System.IO.Path.Combine(candidatePath, PlaywrightFolderName)
            End If

            If Not IsValidPlaywrightDirectory(playwrightDirectory) Then Return

            Dim normalizedRoot As System.String
            Try
                normalizedRoot = System.IO.Path.GetFullPath(root)
            Catch ex As System.Exception
                normalizedRoot = root
            End Try
            If Not seen.Add(normalizedRoot) Then Return
            If IsKnownInvalidLocalCacheRoot(normalizedRoot) Then Return

            Dim runtimeVersionText As System.String = ReadRuntimeVersion(playwrightDirectory, normalizedRoot)
            Dim runtimeVersion As System.Version = ParseVersion(runtimeVersionText)
            If runtimeVersion Is Nothing Then
                Dim folderName As System.String = System.IO.Path.GetFileName(normalizedRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar))
                runtimeVersion = ExtractVersionFromText(folderName)
                If runtimeVersion IsNot Nothing Then runtimeVersionText = runtimeVersion.ToString()
            End If

            result.Add(New PlaywrightRuntimeResolution() With {
                .SourceRoot = normalizedRoot,
                .EffectiveRoot = normalizedRoot,
                .PlaywrightDirectory = playwrightDirectory,
                .BrowsersDirectory = GetOptionalBrowsersDirectory(normalizedRoot),
                .RuntimeVersion = runtimeVersion,
                .RuntimeVersionText = If(runtimeVersionText, System.String.Empty),
                .UsesLocalCache = False
            })
        End Sub

        Private Shared Function IsValidPlaywrightDirectory(playwrightDirectory As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(playwrightDirectory) Then Return False
            Try
                Return System.IO.File.Exists(System.IO.Path.Combine(playwrightDirectory, "node", "win32_x64", "node.exe")) AndAlso
                       System.IO.File.Exists(System.IO.Path.Combine(playwrightDirectory, "package", "cli.js"))
            Catch ex As System.Exception
                Return False
            End Try
        End Function

        Private Shared Function ReadRuntimeVersion(playwrightDirectory As System.String, runtimeRoot As System.String) As System.String
            Dim manifestVersion As System.String = ReadRedInkRuntimeManifestVersion(runtimeRoot)
            If Not System.String.IsNullOrWhiteSpace(manifestVersion) Then Return manifestVersion

            Try
                Dim packageJsonPath As System.String = System.IO.Path.Combine(playwrightDirectory, "package", "package.json")
                If Not System.IO.File.Exists(packageJsonPath) Then Return System.String.Empty
                Dim json As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(packageJsonPath, System.Text.Encoding.UTF8))
                Dim token As Newtonsoft.Json.Linq.JToken = json("version")
                Return If(token Is Nothing, System.String.Empty, token.ToString().Trim())
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function ReadRedInkRuntimeManifestVersion(runtimeRoot As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(runtimeRoot) Then Return System.String.Empty

            Try
                Dim manifestPath As System.String = System.IO.Path.Combine(runtimeRoot, RuntimeManifestFileName)
                If Not System.IO.File.Exists(manifestPath) Then Return System.String.Empty

                Dim json As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(System.IO.File.ReadAllText(manifestPath, System.Text.Encoding.UTF8))
                Dim schemaToken As Newtonsoft.Json.Linq.JToken = json("schema")
                If schemaToken Is Nothing OrElse schemaToken.Type <> Newtonsoft.Json.Linq.JTokenType.Integer OrElse schemaToken.ToObject(Of System.Int32)() <> 1 Then
                    System.Diagnostics.Trace.WriteLine("Ignoring unsupported Red Ink Playwright runtime manifest schema: " & manifestPath)
                    Return System.String.Empty
                End If

                Dim platformToken As Newtonsoft.Json.Linq.JToken = json("platform")
                If platformToken IsNot Nothing Then
                    Dim platform As System.String = platformToken.ToString().Trim()
                    If Not System.String.IsNullOrWhiteSpace(platform) AndAlso Not System.String.Equals(platform, "win-x64", System.StringComparison.OrdinalIgnoreCase) Then
                        System.Diagnostics.Trace.WriteLine("Ignoring Red Ink Playwright runtime manifest for unsupported platform '" & platform & "': " & manifestPath)
                        Return System.String.Empty
                    End If
                End If

                Dim versionToken As Newtonsoft.Json.Linq.JToken = json("playwright_version")
                If versionToken Is Nothing Then Return System.String.Empty
                Return versionToken.ToString().Trim()
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function GetExpectedPlaywrightVersion() As System.Version
            Try
                Dim assembly As System.Reflection.Assembly = GetType(Microsoft.Playwright.Playwright).Assembly

                ' Microsoft.Playwright can expose generic 1.0.x product/informational metadata in
                ' deployed hosts. Such CLR/package metadata must never be treated as the Playwright
                ' driver release. Accept only plausible release versions and continue through all
                ' independent metadata sources until one is found.
                Dim informational As System.Reflection.AssemblyInformationalVersionAttribute =
                    CType(System.Attribute.GetCustomAttribute(assembly, GetType(System.Reflection.AssemblyInformationalVersionAttribute)), System.Reflection.AssemblyInformationalVersionAttribute)
                If informational IsNot Nothing Then
                    Dim parsedInformational As System.Version = ParseVersion(informational.InformationalVersion)
                    If IsUsablePlaywrightReleaseVersion(parsedInformational) Then Return parsedInformational
                End If

                If Not System.String.IsNullOrWhiteSpace(assembly.Location) AndAlso System.IO.File.Exists(assembly.Location) Then
                    Dim fileVersionInfo As System.Diagnostics.FileVersionInfo =
                        System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location)

                    Dim parsedProduct As System.Version = ParseVersion(fileVersionInfo.ProductVersion)
                    If IsUsablePlaywrightReleaseVersion(parsedProduct) Then Return parsedProduct

                    Dim parsedFile As System.Version = ParseVersion(fileVersionInfo.FileVersion)
                    If IsUsablePlaywrightReleaseVersion(parsedFile) Then Return parsedFile
                End If

                ' The loaded assembly identity is normally the most reliable fallback in ClickOnce
                ' deployments because the NuGet package path itself is no longer present there.
                Dim assemblyVersion As System.Version = assembly.GetName().Version
                If IsUsablePlaywrightReleaseVersion(assemblyVersion) Then Return assemblyVersion

                ' Also inspect the reference recorded in SharedLibrary. This follows a NuGet update
                ' automatically at build time and remains available even when file/product metadata
                ' was normalized by deployment packaging.
                Dim referencedVersion As System.Version = GetReferencedPlaywrightVersion()
                If IsUsablePlaywrightReleaseVersion(referencedVersion) Then Return referencedVersion

                ' During development/build output the NuGet version is commonly present in the DLL
                ' path (for example ...\Microsoft.Playwright.1.62.0\lib\...). Keep this only as a
                ' controlled final metadata fallback; it is not required in deployed ClickOnce paths.
                Dim parsedPath As System.Version = ExtractVersionNearPlaywrightPackageName(assembly.Location)
                If IsUsablePlaywrightReleaseVersion(parsedPath) Then Return parsedPath

                Return Nothing
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return Nothing
            End Try
        End Function

        Private Shared Function GetReferencedPlaywrightVersion() As System.Version
            Try
                Dim ownerAssembly As System.Reflection.Assembly = GetType(PlaywrightRuntimeResolver).Assembly
                For Each referencedAssembly As System.Reflection.AssemblyName In ownerAssembly.GetReferencedAssemblies()
                    If referencedAssembly Is Nothing Then Continue For
                    If System.String.Equals(referencedAssembly.Name, "Microsoft.Playwright", System.StringComparison.OrdinalIgnoreCase) Then
                        Return referencedAssembly.Version
                    End If
                Next
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
            Return Nothing
        End Function

        Private Shared Function IsUsablePlaywrightReleaseVersion(version As System.Version) As System.Boolean
            If version Is Nothing Then Return False

            ' 1.0.x is generic compatibility metadata in the affected Playwright deployment and is
            ' not a Playwright NuGet/driver release. In particular, never regress to selecting 1.0.0.
            If version.Major = 1 AndAlso version.Minor = 0 Then Return False

            Return version.Major >= 1 AndAlso version.Minor >= 0
        End Function

        Private Shared Function ExtractVersionNearPlaywrightPackageName(path As System.String) As System.Version
            If System.String.IsNullOrWhiteSpace(path) Then Return Nothing

            Dim match As System.Text.RegularExpressions.Match =
                System.Text.RegularExpressions.Regex.Match(
                    path,
                    "Microsoft\.Playwright[\\\./_-]+(\d+\.\d+(?:\.\d+){0,2})",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase Or
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            If Not match.Success OrElse match.Groups.Count < 2 Then Return Nothing
            Return ParseVersion(match.Groups(1).Value)
        End Function

        Private Shared Function ParseVersion(value As System.String) As System.Version
            If System.String.IsNullOrWhiteSpace(value) Then Return Nothing
            Dim normalized As System.String = value.Trim()
            Dim dashIndex As System.Int32 = normalized.IndexOf("-"c)
            If dashIndex > 0 Then normalized = normalized.Substring(0, dashIndex)
            Dim plusIndex As System.Int32 = normalized.IndexOf("+"c)
            If plusIndex > 0 Then normalized = normalized.Substring(0, plusIndex)
            Dim parsed As System.Version = Nothing
            If System.Version.TryParse(normalized, parsed) Then Return parsed
            Return Nothing
        End Function

        Private Shared Function ExtractVersionFromText(value As System.String) As System.Version
            If System.String.IsNullOrWhiteSpace(value) Then Return Nothing

            Dim match As System.Text.RegularExpressions.Match =
                System.Text.RegularExpressions.Regex.Match(
                    value,
                    "(?<!\d)(\d+\.\d+(?:\.\d+){0,2})(?!\d)",
                    System.Text.RegularExpressions.RegexOptions.CultureInvariant)

            If Not match.Success Then Return Nothing
            Return ParseVersion(match.Groups(1).Value)
        End Function

        Private Shared Function SelectCompatibleCandidate(
            candidates As System.Collections.Generic.List(Of PlaywrightRuntimeResolution),
            expectedVersion As System.Version
        ) As PlaywrightRuntimeResolution
            If candidates Is Nothing OrElse candidates.Count = 0 Then Return Nothing

            Dim compatible As New System.Collections.Generic.List(Of PlaywrightRuntimeResolution)()
            For Each candidate As PlaywrightRuntimeResolution In candidates
                If IsVersionCompatible(expectedVersion, candidate.RuntimeVersion) Then compatible.Add(candidate)
            Next

            If compatible.Count = 0 Then Return Nothing
            compatible.Sort(AddressOf CompareRuntimeCandidatesDescending)
            Return compatible(0)
        End Function

        Private Shared Function IsVersionCompatible(expectedVersion As System.Version, runtimeVersion As System.Version) As System.Boolean
            If expectedVersion Is Nothing Then Return runtimeVersion IsNot Nothing
            If runtimeVersion Is Nothing Then Return False
            Return expectedVersion.Major = runtimeVersion.Major AndAlso expectedVersion.Minor = runtimeVersion.Minor
        End Function

        Private Shared Function CompareRuntimeCandidatesDescending(x As PlaywrightRuntimeResolution, y As PlaywrightRuntimeResolution) As System.Int32
            If x Is Nothing AndAlso y Is Nothing Then Return 0
            If x Is Nothing Then Return 1
            If y Is Nothing Then Return -1
            If x.RuntimeVersion Is Nothing AndAlso y.RuntimeVersion Is Nothing Then Return 0
            If x.RuntimeVersion Is Nothing Then Return 1
            If y.RuntimeVersion Is Nothing Then Return -1
            Return y.RuntimeVersion.CompareTo(x.RuntimeVersion)
        End Function

        Private Shared Function BuildAvailableVersionList(candidates As System.Collections.Generic.List(Of PlaywrightRuntimeResolution)) As System.String
            Dim values As New System.Collections.Generic.List(Of System.String)()
            For Each candidate As PlaywrightRuntimeResolution In candidates
                Dim text As System.String = If(candidate.RuntimeVersion Is Nothing, "unknown", FormatVersion(candidate.RuntimeVersion))
                If Not values.Contains(text) Then values.Add(text)
            Next
            If values.Count = 0 Then Return "none"
            Return System.String.Join(", ", values.ToArray())
        End Function

        Private Shared Function FormatVersion(version As System.Version) As System.String
            If version Is Nothing Then Return "unknown"
            Return version.Major.ToString(System.Globalization.CultureInfo.InvariantCulture) & "." &
                   version.Minor.ToString(System.Globalization.CultureInfo.InvariantCulture) & "." &
                   System.Math.Max(0, version.Build).ToString(System.Globalization.CultureInfo.InvariantCulture)
        End Function

        Private Shared Function GetOptionalBrowsersDirectory(root As System.String) As System.String
            Try
                Dim candidate As System.String = System.IO.Path.Combine(root, "browsers")
                If System.IO.Directory.Exists(candidate) Then Return candidate
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
            Return System.String.Empty
        End Function

        Private Shared Function GetLocalCacheBasePath() As System.String
            Return System.IO.Path.Combine(
                System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                LocalCacheProductFolder,
                LocalCachePlaywrightFolder)
        End Function

        Friend Shared Sub MarkLocalCacheInvalid(runtimeRoot As System.String, reason As System.String)
            Dim normalizedRoot As System.String = NormalizeLocalCacheRoot(runtimeRoot)
            If System.String.IsNullOrWhiteSpace(normalizedRoot) Then Return

            SyncLock InvalidLocalCacheLock
                InvalidLocalCacheRoots.Add(normalizedRoot)
            End SyncLock

            Try
                If System.IO.Directory.Exists(normalizedRoot) Then
                    System.IO.File.WriteAllText(
                        System.IO.Path.Combine(normalizedRoot, InvalidRuntimeMarkerFileName),
                        "invalid",
                        System.Text.Encoding.UTF8)
                End If
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try

            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog(
                "[ToolAvailability] Playwright local cache marked invalid; target=" & CompactLogValue(normalizedRoot) &
                "; reason=" & CompactLogValue(reason))
        End Sub

        Private Shared Function NormalizeLocalCacheRoot(runtimeRoot As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(runtimeRoot) Then Return System.String.Empty

            Try
                Dim normalizedRoot As System.String = System.IO.Path.GetFullPath(runtimeRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar))
                Dim cacheBase As System.String = System.IO.Path.GetFullPath(GetLocalCacheBasePath()).TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                Dim requiredPrefix As System.String = cacheBase & System.IO.Path.DirectorySeparatorChar
                If Not normalizedRoot.StartsWith(requiredPrefix, System.StringComparison.OrdinalIgnoreCase) Then Return System.String.Empty
                Return normalizedRoot
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return System.String.Empty
            End Try
        End Function

        Private Shared Function IsKnownInvalidLocalCacheRoot(runtimeRoot As System.String) As System.Boolean
            Dim normalizedRoot As System.String = NormalizeLocalCacheRoot(runtimeRoot)
            If System.String.IsNullOrWhiteSpace(normalizedRoot) Then Return False

            SyncLock InvalidLocalCacheLock
                If InvalidLocalCacheRoots.Contains(normalizedRoot) Then Return True
            End SyncLock

            Try
                Return System.IO.File.Exists(System.IO.Path.Combine(normalizedRoot, InvalidRuntimeMarkerFileName))
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                Return False
            End Try
        End Function

        Private Shared Sub ClearLocalCacheInvalid(runtimeRoot As System.String)
            Dim normalizedRoot As System.String = NormalizeLocalCacheRoot(runtimeRoot)
            If System.String.IsNullOrWhiteSpace(normalizedRoot) Then Return

            SyncLock InvalidLocalCacheLock
                InvalidLocalCacheRoots.Remove(normalizedRoot)
            End SyncLock

            Try
                Dim markerPath As System.String = System.IO.Path.Combine(normalizedRoot, InvalidRuntimeMarkerFileName)
                If System.IO.File.Exists(markerPath) Then System.IO.File.Delete(markerPath)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
        End Sub

        Private Shared Sub StartLocalCachePreparation(source As PlaywrightRuntimeResolution, cachePreparationStarted As System.Action)
            If source Is Nothing OrElse source.RuntimeVersion Is Nothing Then Return

            Dim key As System.String = FormatVersion(source.RuntimeVersion)
            SyncLock BackgroundCacheLock
                Dim existing As System.Threading.Tasks.Task = Nothing
                If BackgroundCacheTasks.TryGetValue(key, existing) AndAlso existing IsNot Nothing AndAlso Not existing.IsCompleted Then Return

                If cachePreparationStarted IsNot Nothing Then
                    Try
                        cachePreparationStarted.Invoke()
                    Catch ex As System.Exception
                        System.Diagnostics.Trace.WriteLine("Playwright background cache notification failed: " & ex.ToString())
                    End Try
                End If

                Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog("[ToolAvailability] Playwright cache provisioning started; version=" & key & "; source=" & CompactLogValue(source.SourceRoot))

                Dim task As System.Threading.Tasks.Task = System.Threading.Tasks.Task.Run(
                    Sub()
                        Try
                            Dim cached As PlaywrightRuntimeResolution = MaterializeLocalCache(source, Nothing)
                            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog("[ToolAvailability] Playwright cache provisioning completed; version=" & key & "; target=" & CompactLogValue(cached.EffectiveRoot))
                        Catch ex As System.Exception
                            Global.SharedLibrary.SharedLibrary.UpdateHandler.WriteUpdateLog("[ToolAvailability] Playwright cache provisioning failed; version=" & key & "; error=" & CompactLogValue(ex.Message))
                        Finally
                            SyncLock BackgroundCacheLock
                                BackgroundCacheTasks.Remove(key)
                            End SyncLock
                        End Try
                    End Sub)
                BackgroundCacheTasks(key) = task
            End SyncLock
        End Sub

        Private Shared Function CompactLogValue(value As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(value) Then Return "(none)"
            Return value.Replace(System.Environment.NewLine, " ").Replace(System.Convert.ToChar(13), " "c).Replace(System.Convert.ToChar(10), " "c).Trim()
        End Function

        Private Shared Function MaterializeLocalCache(source As PlaywrightRuntimeResolution, cachePreparationStarted As System.Action) As PlaywrightRuntimeResolution
            If source Is Nothing Then Throw New System.ArgumentNullException(NameOf(source))
            Dim versionFolder As System.String = If(source.RuntimeVersion Is Nothing, "unknown", FormatVersion(source.RuntimeVersion))
            Dim localRoot As System.String = System.IO.Path.Combine(GetLocalCacheBasePath(), versionFolder)
            Dim localPlaywright As System.String = System.IO.Path.Combine(localRoot, PlaywrightFolderName)
            Dim mutexName As System.String = "RedInk.PlaywrightRuntimeCache." & versionFolder.Replace(".", "_")

            Using cacheMutex As New System.Threading.Mutex(False, mutexName)
                Dim lockTaken As System.Boolean = False
                Try
                    Try
                        lockTaken = cacheMutex.WaitOne(System.TimeSpan.FromSeconds(60))
                    Catch ex As System.Threading.AbandonedMutexException
                        lockTaken = True
                    End Try
                    If Not lockTaken Then Throw New System.TimeoutException("Timed out waiting for another Red Ink process to finish preparing the Playwright runtime cache.")

                    Dim sourceHasBrowsers As System.Boolean = Not System.String.IsNullOrWhiteSpace(source.BrowsersDirectory) AndAlso System.IO.Directory.Exists(source.BrowsersDirectory)
                    Dim localBrowsers As System.String = System.IO.Path.Combine(localRoot, "browsers")
                    Dim localVersion As System.Version = ParseVersion(ReadRuntimeVersion(localPlaywright, localRoot))
                    Dim cacheVersionMatches As System.Boolean = IsVersionCompatible(source.RuntimeVersion, localVersion)
                    Dim cacheWasInvalid As System.Boolean = IsKnownInvalidLocalCacheRoot(localRoot)
                    Dim needsRefresh As System.Boolean = cacheWasInvalid OrElse Not IsValidPlaywrightDirectory(localPlaywright) OrElse Not cacheVersionMatches OrElse (sourceHasBrowsers AndAlso Not System.IO.Directory.Exists(localBrowsers))

                    If needsRefresh Then
                        If cachePreparationStarted IsNot Nothing Then cachePreparationStarted.Invoke()
                        Dim stagingRoot As System.String = localRoot & ".staging-" & System.Guid.NewGuid().ToString("N")
                        Try
                            If System.IO.Directory.Exists(stagingRoot) Then System.IO.Directory.Delete(stagingRoot, True)
                            System.IO.Directory.CreateDirectory(stagingRoot)
                            CopyDirectory(source.PlaywrightDirectory, System.IO.Path.Combine(stagingRoot, PlaywrightFolderName))
                            CopyRuntimeManifestIfPresent(source.SourceRoot, stagingRoot)
                            If sourceHasBrowsers Then
                                CopyDirectory(source.BrowsersDirectory, System.IO.Path.Combine(stagingRoot, "browsers"))
                            End If
                            If System.IO.Directory.Exists(localRoot) Then System.IO.Directory.Delete(localRoot, True)
                            System.IO.Directory.Move(stagingRoot, localRoot)
                            ClearLocalCacheInvalid(localRoot)
                        Finally
                            If System.IO.Directory.Exists(stagingRoot) Then
                                Try
                                    System.IO.Directory.Delete(stagingRoot, True)
                                Catch ex As System.Exception
                                    System.Diagnostics.Trace.WriteLine(ex.ToString())
                                End Try
                            End If
                        End Try
                    End If
                Finally
                    If lockTaken Then
                        Try
                            cacheMutex.ReleaseMutex()
                        Catch ex As System.Exception
                            System.Diagnostics.Trace.WriteLine(ex.ToString())
                        End Try
                    End If
                End Try
            End Using

            Return New PlaywrightRuntimeResolution() With {
                .SourceRoot = source.SourceRoot,
                .EffectiveRoot = localRoot,
                .PlaywrightDirectory = localPlaywright,
                .BrowsersDirectory = GetOptionalBrowsersDirectory(localRoot),
                .RuntimeVersion = source.RuntimeVersion,
                .RuntimeVersionText = source.RuntimeVersionText,
                .UsesLocalCache = True
            }
        End Function

        Private Shared Sub CopyRuntimeManifestIfPresent(sourceRoot As System.String, targetRoot As System.String)
            If System.String.IsNullOrWhiteSpace(sourceRoot) OrElse System.String.IsNullOrWhiteSpace(targetRoot) Then Return

            Try
                Dim sourceManifest As System.String = System.IO.Path.Combine(sourceRoot, RuntimeManifestFileName)
                If Not System.IO.File.Exists(sourceManifest) Then Return
                System.IO.File.Copy(sourceManifest, System.IO.Path.Combine(targetRoot, RuntimeManifestFileName), True)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
            End Try
        End Sub

        Private Shared Sub CopyDirectory(sourceDirectory As System.String, targetDirectory As System.String)
            System.IO.Directory.CreateDirectory(targetDirectory)
            For Each filePath As System.String In System.IO.Directory.GetFiles(sourceDirectory)
                Dim targetFile As System.String = System.IO.Path.Combine(targetDirectory, System.IO.Path.GetFileName(filePath))
                System.IO.File.Copy(filePath, targetFile, True)
            Next
            For Each directoryPath As System.String In System.IO.Directory.GetDirectories(sourceDirectory)
                Dim targetChild As System.String = System.IO.Path.Combine(targetDirectory, System.IO.Path.GetFileName(directoryPath))
                CopyDirectory(directoryPath, targetChild)
            Next
        End Sub
    End Class
End Namespace
