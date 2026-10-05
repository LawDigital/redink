' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
' Host-owned authorization and physical source/artifact containment. Unknown denies.
Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class SemanticArchiveAccessContext
        Private ReadOnly _check As System.Func(Of System.String, SemanticArchiveAccessDecision)
        Private ReadOnly _denialCode As System.String
        Private ReadOnly _denialMessage As System.String
        Private ReadOnly _validUntilUtc As System.DateTimeOffset?
        Private ReadOnly _validityClock As System.Diagnostics.Stopwatch
        Private ReadOnly _validityDuration As System.TimeSpan
        Private _authorizationExpired As System.Int32
        Private _sourceChecks As System.Int64
        Public ReadOnly Property PrincipalId As System.String

        Public ReadOnly Property SourceChecks As System.Int64
            Get
                Return System.Threading.Interlocked.Read(_sourceChecks)
            End Get
        End Property

        Public Sub New(principalId As System.String, sourceAccessCheck As System.Func(Of System.String, SemanticArchiveAccessDecision),
                       Optional denialCode As System.String = "", Optional denialMessage As System.String = "",
                       Optional validUntilUtc As System.DateTimeOffset? = Nothing)
            Me.PrincipalId = If(principalId, "")
            _check = sourceAccessCheck
            _denialCode = If(denialCode, "")
            _denialMessage = If(denialMessage, "")
            _validUntilUtc = validUntilUtc
            If validUntilUtc.HasValue Then
                _validityDuration = validUntilUtc.Value - System.DateTimeOffset.UtcNow
                _validityClock = System.Diagnostics.Stopwatch.StartNew()
            End If
        End Sub

        Private Function HasExpiredAuthorization() As System.Boolean
            If Not _validUntilUtc.HasValue Then Return False
            If System.Threading.Volatile.Read(_authorizationExpired) <> 0 Then Return True
            If System.DateTimeOffset.UtcNow >= _validUntilUtc.Value OrElse _validityClock.Elapsed >= _validityDuration Then
                System.Threading.Interlocked.Exchange(_authorizationExpired, 1)
                Return True
            End If
            Return False
        End Function

        Public ReadOnly Property DenialCode As System.String
            Get
                If _denialCode.Length > 0 Then Return _denialCode
                If HasExpiredAuthorization() Then Return "requester_identity_unverified"
                Return ""
            End Get
        End Property

        Public ReadOnly Property DenialMessage As System.String
            Get
                If _denialCode.Length > 0 Then Return _denialMessage
                If HasExpiredAuthorization() Then
                    Return "The independently verified requester authorization expired. Semantic Archive access requires fresh verification for this request."
                End If
                Return ""
            End Get
        End Property

        ''' <summary>
        ''' Use only for a directly interacting local user. Remote/delegated hosts must
        ''' supply the requesting principal's independently verified source policy.
        ''' </summary>
        Public Shared Function CreateForCurrentUser() As SemanticArchiveAccessContext
            Dim identity As System.String = ""
            Try
                Using current As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
                    If current.User IsNot Nothing Then identity = current.User.Value
                End Using
            Catch ex As System.Exception
                Return New SemanticArchiveAccessContext("", Nothing)
            End Try
            Dim capturedIdentity As System.String = identity
            Return New SemanticArchiveAccessContext(identity,
                Function(path As System.String) As SemanticArchiveAccessDecision
                    Try
                        Using current As System.Security.Principal.WindowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent()
                            If current.User Is Nothing OrElse Not System.String.Equals(current.User.Value, capturedIdentity, System.StringComparison.Ordinal) Then Return SemanticArchiveAccessDecision.Denied
                        End Using
                        Using stream As New System.IO.FileStream(path, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                            If Not stream.CanRead Then Return SemanticArchiveAccessDecision.Denied
                        End Using
                        Return SemanticArchiveAccessDecision.Allowed
                    Catch ex As System.Exception
                        Return SemanticArchiveAccessDecision.Denied
                    End Try
                End Function)
        End Function

        Public Function CanReadSource(path As System.String) As System.Boolean
            If DenialCode.Length > 0 Then Return False
            If System.String.IsNullOrWhiteSpace(PrincipalId) OrElse _check Is Nothing OrElse System.String.IsNullOrWhiteSpace(path) Then Return False
            System.Threading.Interlocked.Increment(_sourceChecks)
            Try
                Return _check(path) = SemanticArchiveAccessDecision.Allowed AndAlso DenialCode.Length = 0
            Catch ex As System.Exception
                Return False
            End Try
        End Function
    End Class

    Public NotInheritable Class SemanticArchivePathGuard
        Private Sub New()
        End Sub

        Private Shared ReadOnly Property PathComparison As System.StringComparison
            Get
                Return If(System.Environment.OSVersion.Platform = System.PlatformID.Win32NT, System.StringComparison.OrdinalIgnoreCase, System.StringComparison.Ordinal)
            End Get
        End Property

        Public Shared Function CanonicalPath(path As System.String) As System.String
            If System.String.IsNullOrWhiteSpace(path) Then Throw New System.ArgumentException("A filesystem path is required.", NameOf(path))
            Dim raw As System.String = path.Trim().Trim(Microsoft.VisualBasic.ChrW(34))
            If raw.StartsWith("\\?\", System.StringComparison.Ordinal) OrElse raw.StartsWith("\\.\", System.StringComparison.Ordinal) Then Throw New System.UnauthorizedAccessException("Device paths are not archive paths.")
            Dim expanded As System.String = SharedMethods.ExpandEnvironmentVariables(raw)
            If System.String.IsNullOrWhiteSpace(expanded) Then Throw New System.ArgumentException("The archive path could not be expanded.", NameOf(path))
            If System.Text.RegularExpressions.Regex.IsMatch(expanded, "%[^%]+%", System.Text.RegularExpressions.RegexOptions.CultureInvariant, System.TimeSpan.FromMilliseconds(100)) Then
                Throw New System.ArgumentException("The archive path contains an unresolved environment variable or placeholder.", NameOf(path))
            End If
            If expanded.StartsWith("\\?\", System.StringComparison.Ordinal) OrElse expanded.StartsWith("\\.\", System.StringComparison.Ordinal) Then Throw New System.UnauthorizedAccessException("Device paths are not archive paths.")
            Dim full As System.String = System.IO.Path.GetFullPath(expanded)
            Dim root As System.String = System.IO.Path.GetPathRoot(full)
            If full.Substring(root.Length).IndexOf(":"c) >= 0 Then Throw New System.UnauthorizedAccessException("Alternate data streams are not archive paths.")
            Return If(full.Length = root.Length, full, full.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar))
        End Function

        Public Shared Function RequireWindowsSourcePath(path As System.String) As System.String
            Dim full As System.String = CanonicalPath(path)
            If full.Length > 259 Then Throw New System.IO.PathTooLongException("The original path exceeds the supported Windows source path limit.")
            For Each component As System.String In full.Split(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, System.StringSplitOptions.RemoveEmptyEntries)
                If component.Length > 255 Then Throw New System.IO.PathTooLongException("An original path component exceeds the Windows source path limit.")
            Next
            Return full
        End Function

        Public Shared Function RequireWindowsCompatiblePath(path As System.String, Optional reservedSuffixLength As System.Int32 = 0) As System.String
            If reservedSuffixLength < 0 OrElse reservedSuffixLength > 240 Then Throw New System.ArgumentOutOfRangeException(NameOf(reservedSuffixLength))
            Dim full As System.String = CanonicalPath(path)
            If full.Length + reservedSuffixLength > 240 Then Throw New System.IO.PathTooLongException("The complete generated path exceeds the supported Windows path budget.")
            For Each component As System.String In full.Split(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, System.StringSplitOptions.RemoveEmptyEntries)
                If component.Length > 240 Then Throw New System.IO.PathTooLongException("A path component exceeds the supported Windows path budget.")
            Next
            Return full
        End Function

        Public Shared Function IsContainedPath(root As System.String, candidate As System.String) As System.Boolean
            Dim normalizedRoot As System.String = CanonicalPath(root)
            Dim normalizedCandidate As System.String = CanonicalPath(candidate)
            Return System.String.Equals(normalizedRoot, normalizedCandidate, PathComparison) OrElse normalizedCandidate.StartsWith(normalizedRoot.TrimEnd(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar) & System.IO.Path.DirectorySeparatorChar, PathComparison)
        End Function

        ''' <summary>
        ''' Reject every reparse component, including ancestors of the registered
        ''' root. This deliberately does not treat lexical normalization as physical
        ''' containment. Unreadable attributes are an authorization failure.
        ''' </summary>
        Public Shared Function ValidateContainedPath(root As System.String, path As System.String, mustExist As System.Boolean) As System.String
            Dim normalizedRoot As System.String = CanonicalPath(root)
            Dim full As System.String = CanonicalPath(path)
            If Not IsContainedPath(normalizedRoot, full) Then Throw New System.UnauthorizedAccessException("The path is outside its registered archive root.")
            ValidatePhysicalComponents(normalizedRoot, mustExist)
            ValidatePhysicalComponents(full, mustExist)
            Return full
        End Function

        Private Shared Sub ValidatePhysicalComponents(path As System.String, mustExist As System.Boolean)
            Dim volume As System.String = System.IO.Path.GetPathRoot(path)
            Dim current As System.String = volume
            Dim components As System.String() = path.Substring(volume.Length).Split(New System.Char() {System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar}, System.StringSplitOptions.RemoveEmptyEntries)
            CheckComponent(current, mustExist)
            For Each component As System.String In components
                current = System.IO.Path.Combine(current, component)
                If Not CheckComponent(current, mustExist) Then Return
            Next
        End Sub

        Private Shared Function CheckComponent(path As System.String, mustExist As System.Boolean) As System.Boolean
            Try
                Dim attributes As System.IO.FileAttributes = System.IO.File.GetAttributes(path)
                If (attributes And System.IO.FileAttributes.ReparsePoint) <> 0 Then Throw New System.UnauthorizedAccessException("Archive traversal through a reparse point is forbidden.")
                Return True
            Catch ex As System.IO.FileNotFoundException
                If mustExist Then Throw
                Return False
            Catch ex As System.IO.DirectoryNotFoundException
                If mustExist Then Throw
                Return False
            End Try
        End Function

        ''' <summary>Attributes-only identity for restrictive maintenance; never a content-read authorization.</summary>
        Public Shared Function GetVerifiedSourceIdentityForMaintenance(root As System.String, path As System.String) As System.String
            Dim full As System.String = ValidateContainedPath(root, path, True)
            Dim physicalRoot As System.String = GetVerifiedDirectoryIdentity(root)
            Dim relative As System.String = full.Substring(CanonicalPath(root).Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
            If relative.Length = 0 Then Throw New System.UnauthorizedAccessException("Maintenance requires a file below its registered root.")
            Using handle As Microsoft.Win32.SafeHandles.SafeFileHandle = CreateFileForDirectory(full, &H80UI, &H3UI, System.IntPtr.Zero, 3UI, &H2200000UI, System.IntPtr.Zero)
                If handle Is Nothing OrElse handle.IsInvalid Then Throw New System.UnauthorizedAccessException("The source identity is unavailable for restrictive maintenance.", New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()))
                Dim information As NativeFileInformation
                If Not GetFileInformationByHandle(handle, information) OrElse (information.Attributes And (CUInt(System.IO.FileAttributes.Directory) Or CUInt(System.IO.FileAttributes.ReparsePoint))) <> 0UI Then Throw New System.UnauthorizedAccessException("The maintenance source is not an ordinary file.")
                Dim actual As System.String = GetPhysicalHandlePath(handle).ToUpperInvariant()
                If Not System.String.Equals(actual, CanonicalPath(System.IO.Path.Combine(physicalRoot, relative)), System.StringComparison.OrdinalIgnoreCase) Then Throw New System.UnauthorizedAccessException("The maintenance source no longer matches its registered physical scope.")
                Return actual
            End Using
        End Function

        Public Shared Function GetVerifiedDirectoryIdentity(path As System.String) As System.String
            Dim full As System.String = ValidateContainedPath(path, path, True)
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then Throw New System.PlatformNotSupportedException("Physical Windows directory identity is required.")
            Using handle As Microsoft.Win32.SafeHandles.SafeFileHandle = CreateFileForDirectory(full, &H80UI, &H3UI, System.IntPtr.Zero, 3UI, &H2200000UI, System.IntPtr.Zero)
                If handle Is Nothing OrElse handle.IsInvalid Then Throw New System.UnauthorizedAccessException("The artifact protection domain could not be verified.", New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()))
                Dim information As NativeFileInformation
                If Not GetFileInformationByHandle(handle, information) OrElse (information.Attributes And CUInt(System.IO.FileAttributes.Directory)) = 0UI OrElse (information.Attributes And CUInt(System.IO.FileAttributes.ReparsePoint)) <> 0UI Then Throw New System.UnauthorizedAccessException("The artifact parent is not an ordinary physical directory.")
                Return GetPhysicalHandlePath(handle).ToUpperInvariant()
            End Using
        End Function

        Public Shared Function OpenContainedRead(root As System.String, path As System.String) As System.IO.FileStream
            Dim physicalPath As System.String = Nothing
            Return OpenContainedReadCore(root, path, physicalPath)
        End Function

        ''' <summary>
        ''' A deduplication identity obtained from a physically contained file handle.
        ''' The logical binding still controls authorization and the saved source path.
        ''' Final names are used, so distinct hard-link names remain distinct sources.
        ''' </summary>
        Public Shared Function GetVerifiedSourceIdentity(root As System.String, path As System.String) As System.String
            Dim physicalPath As System.String = Nothing
            Using stream As System.IO.FileStream = OpenContainedReadCore(root, path, physicalPath)
                Return If(System.Environment.OSVersion.Platform = System.PlatformID.Win32NT, physicalPath.ToUpperInvariant(), physicalPath)
            End Using
        End Function

        Private Shared Function OpenContainedReadCore(root As System.String, path As System.String, ByRef verifiedPhysicalPath As System.String) As System.IO.FileStream
            verifiedPhysicalPath = Nothing
            Dim normalizedRoot As System.String = CanonicalPath(root)
            Dim full As System.String = ValidateContainedPath(normalizedRoot, path, True)
            If System.Environment.OSVersion.Platform <> System.PlatformID.Win32NT Then
                Dim portableStream As New System.IO.FileStream(full, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Try
                    ValidateContainedPath(normalizedRoot, full, True)
                    verifiedPhysicalPath = full
                    Return portableStream
                Catch ex As System.Exception
                    portableStream.Dispose()
                    Throw
                End Try
            End If

            ' Both sides of the comparison come from real handles. A mapped drive
            ' may resolve to a UNC name; comparing that with its logical drive letter
            ' would reject a legitimate source. Keep the original lexical scope and
            ' all reparse checks, then compare within this opened physical root.
            Using rootHandle As Microsoft.Win32.SafeHandles.SafeFileHandle = CreateFileForDirectory(
                normalizedRoot, &H80UI, &H3UI, System.IntPtr.Zero, 3UI,
                &H2200000UI, System.IntPtr.Zero)
                If rootHandle Is Nothing OrElse rootHandle.IsInvalid Then
                    Throw New System.UnauthorizedAccessException("The registered archive root could not be opened for physical validation.", New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()))
                End If
                Dim rootInformation As NativeFileInformation
                If Not GetFileInformationByHandle(rootHandle, rootInformation) OrElse
                    (rootInformation.Attributes And CUInt(System.IO.FileAttributes.Directory)) = 0UI OrElse
                    (rootInformation.Attributes And CUInt(System.IO.FileAttributes.ReparsePoint)) <> 0UI Then
                    Throw New System.UnauthorizedAccessException("The registered physical root is not a verifiable ordinary directory.")
                End If
                Dim physicalRoot As System.String = GetPhysicalHandlePath(rootHandle)
                Dim relative As System.String = full.Substring(normalizedRoot.Length).TrimStart(System.IO.Path.DirectorySeparatorChar, System.IO.Path.AltDirectorySeparatorChar)
                If relative.Length = 0 Then Throw New System.UnauthorizedAccessException("Archive reads require a file below the registered root.")
                Dim expectedPhysicalPath As System.String = CanonicalPath(System.IO.Path.Combine(physicalRoot, relative))
                Dim stream As New System.IO.FileStream(full, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read)
                Try
                    ValidateContainedPath(normalizedRoot, full, True)
                    Dim actual As System.String = GetPhysicalHandlePath(stream.SafeFileHandle)
                    Dim currentPhysicalRoot As System.String = GetPhysicalHandlePath(rootHandle)
                    If Not System.String.Equals(physicalRoot, currentPhysicalRoot, PathComparison) OrElse
                        Not IsContainedPath(physicalRoot, actual) OrElse
                        Not System.String.Equals(actual, expectedPhysicalPath, PathComparison) Then
                        Throw New System.UnauthorizedAccessException("The opened source does not match its registered physical root and relative path.")
                    End If
                    verifiedPhysicalPath = actual
                    Return stream
                Catch ex As System.Exception
                    stream.Dispose()
                    Throw
                End Try
            End Using
        End Function

        Private Shared Function GetPhysicalHandlePath(handle As Microsoft.Win32.SafeHandles.SafeFileHandle) As System.String
            Dim buffer As New System.Text.StringBuilder(32768)
            Dim length As System.UInt32 = GetFinalPathNameByHandle(handle, buffer, CUInt(buffer.Capacity), 0UI)
            If length = 0UI OrElse length >= CUInt(buffer.Capacity) Then Throw New System.UnauthorizedAccessException("The physical filesystem location could not be established.", New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error()))
            Dim actual As System.String = buffer.ToString()
            If actual.StartsWith("\\?\UNC\", System.StringComparison.OrdinalIgnoreCase) Then
                actual = "\\" & actual.Substring(8)
            ElseIf actual.StartsWith("\\?\", System.StringComparison.Ordinal) Then
                actual = actual.Substring(4)
            End If
            If Not System.IO.Path.IsPathRooted(actual) Then Throw New System.UnauthorizedAccessException("The filesystem returned an unsupported physical path name.")
            Return CanonicalPath(actual)
        End Function

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure NativeFileInformation
            Public Attributes As System.UInt32
            Public CreationTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public LastAccessTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public LastWriteTime As System.Runtime.InteropServices.ComTypes.FILETIME
            Public VolumeSerialNumber As System.UInt32
            Public FileSizeHigh As System.UInt32
            Public FileSizeLow As System.UInt32
            Public NumberOfLinks As System.UInt32
            Public FileIndexHigh As System.UInt32
            Public FileIndexLow As System.UInt32
        End Structure

        <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="CreateFileW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function CreateFileForDirectory(path As System.String, desiredAccess As System.UInt32, shareMode As System.UInt32,
                                                      securityAttributes As System.IntPtr, creationDisposition As System.UInt32,
                                                      flagsAndAttributes As System.UInt32, templateHandle As System.IntPtr) As Microsoft.Win32.SafeHandles.SafeFileHandle
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function GetFileInformationByHandle(handle As Microsoft.Win32.SafeHandles.SafeFileHandle, ByRef information As NativeFileInformation) As System.Boolean
        End Function

        <System.Runtime.InteropServices.DllImport("kernel32.dll", EntryPoint:="GetFinalPathNameByHandleW", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, ExactSpelling:=True, SetLastError:=True)>
        Private Shared Function GetFinalPathNameByHandle(handle As Microsoft.Win32.SafeHandles.SafeFileHandle, path As System.Text.StringBuilder, capacity As System.UInt32, flags As System.UInt32) As System.UInt32
        End Function
    End Class
End Namespace
