' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: Crashlogger.vb
' Purpose:
'   Best-effort, low-overhead crash/diagnostics logger for the Office add-in hosts.
'   Installs process-wide exception handlers only when explicitly enabled, so that
'   users for whom diagnostics are switched off pay no runtime cost.
'
' Notes:
'   - Enabling is driven per add-in by My.Settings.CrashLog, which is reconciled from
'     the INI parameter INI_Crashlog after configuration load (takes effect on the
'     next host launch, because the INI is read after ThisAddIn_Startup).
'   - The logger must never throw into the host; all paths fail silently.
'   - On-disk retention: the active log rolls to ri-crashlog.previous.txt at 10 MB,
'     giving roughly 20 MB of headroom (several days of typical collection).
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary

    Public NotInheritable Class RiCrashLogger

        Private Sub New()
        End Sub

        Private Const MaximumLogFileSize As System.Int64 =
            10L * 1024L * 1024L

        Private Const FileMutexName As System.String =
            "Local\RedInk_RiCrashLog_Mutex"

        Private Shared ReadOnly InitializationLock As New System.Object()

        Private Shared DiagnosticsEnabled As System.Boolean
        Private Shared Initialized As System.Boolean

        Private Shared SessionId As System.String
        Private Shared AddInName As System.String
        Private Shared RdvVersionValue As System.String
        Private Shared LogFilePathValue As System.String

        Private Shared AddInAssembly As System.Reflection.Assembly
        Private Shared SnapshotTimer As System.Threading.Timer

        Public Shared ReadOnly Property LogFilePath As System.String
            Get
                Return LogFilePathValue
            End Get
        End Property

        Public Shared Sub Initialize(
            ByVal addInDisplayName As System.String,
            ByVal executingAddInAssembly As System.Reflection.Assembly,
            ByVal diagnosticsAreEnabled As System.Boolean,
            Optional ByVal captureExtendedSnapshot As System.Boolean = True,
            Optional ByVal rdvVersion As System.String = Nothing)

            If Not diagnosticsAreEnabled Then
                Return
            End If

            SyncLock InitializationLock

                If Initialized Then
                    Return
                End If

                DiagnosticsEnabled = True
                Initialized = True

                SessionId =
                    System.Guid.NewGuid().ToString("N")

                AddInName =
                    If(
                        System.String.IsNullOrWhiteSpace(addInDisplayName),
                        "Unknown add-in",
                        addInDisplayName)

                RdvVersionValue = rdvVersion
                AddInAssembly = executingAddInAssembly

                Dim appDataDirectory As System.String =
                    System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.ApplicationData)

                Dim redInkDirectory As System.String =
                    System.IO.Path.Combine(
                        appDataDirectory,
                        "redink")

                LogFilePathValue =
                    System.IO.Path.Combine(
                        redInkDirectory,
                        "RI-CrashLog.txt")

                Try
                    System.IO.Directory.CreateDirectory(redInkDirectory)
                Catch ex As System.Exception
                    DiagnosticsEnabled = False
                    Return
                End Try

                WriteCollectorReadme(redInkDirectory)

                RegisterExceptionHandlers()

            End SyncLock

            AppendRecord(
                "SESSION_START",
                BuildStartupInformation())

            If captureExtendedSnapshot Then

                Try
                    SnapshotTimer =
                        New System.Threading.Timer(
                            AddressOf DeferredSnapshotTimerCallback,
                            Nothing,
                            8000,
                            System.Threading.Timeout.Infinite)

                Catch ex As System.Exception
                    AppendRecord(
                        "SNAPSHOT_TIMER_ERROR",
                        FormatException(ex))
                End Try

            End If

        End Sub

        Public Shared Sub Shutdown(
            Optional ByVal reason As System.String = "Add-in shutdown")

            If Not DiagnosticsEnabled Then
                Return
            End If

            AppendRecord(
                "SESSION_END",
                reason)

            SyncLock InitializationLock

                DiagnosticsEnabled = False

                UnregisterExceptionHandlers()

                Dim timerToDispose As System.Threading.Timer =
                    SnapshotTimer

                SnapshotTimer = Nothing

                If timerToDispose IsNot Nothing Then

                    Try
                        timerToDispose.Dispose()
                    Catch ex As System.Exception
                    End Try

                End If

            End SyncLock

        End Sub

        Public Shared Sub LogMarker(
            ByVal eventName As System.String,
            Optional ByVal details As System.String = Nothing)

            AppendRecord(
                eventName,
                details)

        End Sub

        Public Shared Sub LogException(
            ByVal context As System.String,
            ByVal ex As System.Exception)

            AppendRecord(
                context,
                FormatException(ex))

        End Sub

        Private Shared Sub RegisterExceptionHandlers()

            AddHandler System.AppDomain.CurrentDomain.UnhandledException,
                AddressOf CurrentDomain_UnhandledException

            AddHandler System.AppDomain.CurrentDomain.AssemblyResolve,
                AddressOf CurrentDomain_AssemblyResolve

            AddHandler System.AppDomain.CurrentDomain.ProcessExit,
                AddressOf CurrentDomain_ProcessExit

            AddHandler System.AppDomain.CurrentDomain.DomainUnload,
                AddressOf CurrentDomain_DomainUnload

            AddHandler System.Windows.Forms.Application.ThreadException,
                AddressOf WindowsForms_ThreadException

            AddHandler System.Threading.Tasks.TaskScheduler.UnobservedTaskException,
                AddressOf TaskScheduler_UnobservedTaskException

        End Sub

        Private Shared Sub UnregisterExceptionHandlers()

            Try
                RemoveHandler System.AppDomain.CurrentDomain.UnhandledException,
                    AddressOf CurrentDomain_UnhandledException
            Catch ex As System.Exception
            End Try

            Try
                RemoveHandler System.AppDomain.CurrentDomain.AssemblyResolve,
                    AddressOf CurrentDomain_AssemblyResolve
            Catch ex As System.Exception
            End Try

            Try
                RemoveHandler System.AppDomain.CurrentDomain.ProcessExit,
                    AddressOf CurrentDomain_ProcessExit
            Catch ex As System.Exception
            End Try

            Try
                RemoveHandler System.AppDomain.CurrentDomain.DomainUnload,
                    AddressOf CurrentDomain_DomainUnload
            Catch ex As System.Exception
            End Try

            Try
                RemoveHandler System.Windows.Forms.Application.ThreadException,
                    AddressOf WindowsForms_ThreadException
            Catch ex As System.Exception
            End Try

            Try
                RemoveHandler System.Threading.Tasks.TaskScheduler.UnobservedTaskException,
                    AddressOf TaskScheduler_UnobservedTaskException
            Catch ex As System.Exception
            End Try

        End Sub

        Private Shared Sub CurrentDomain_UnhandledException(
            ByVal sender As System.Object,
            ByVal e As System.UnhandledExceptionEventArgs)

            Dim ex As System.Exception =
                TryCast(
                    e.ExceptionObject,
                    System.Exception)

            Dim details As New System.Text.StringBuilder()

            details.AppendLine(
                "IsTerminating=" &
                e.IsTerminating.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            If ex IsNot Nothing Then
                details.AppendLine(FormatException(ex))
            Else
                details.AppendLine(
                    "ExceptionObject=" &
                    SafeToString(e.ExceptionObject))
            End If

            AppendRecord(
                "APPDOMAIN_UNHANDLED_EXCEPTION",
                details.ToString())

        End Sub

        Private Shared Sub WindowsForms_ThreadException(
            ByVal sender As System.Object,
            ByVal e As System.Threading.ThreadExceptionEventArgs)

            AppendRecord(
                "WINDOWS_FORMS_THREAD_EXCEPTION",
                FormatException(e.Exception))

        End Sub

        Private Shared Sub TaskScheduler_UnobservedTaskException(
            ByVal sender As System.Object,
            ByVal e As System.Threading.Tasks.UnobservedTaskExceptionEventArgs)

            Dim details As New System.Text.StringBuilder()

            details.AppendLine(
                "Observed=" &
                e.Observed.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            details.AppendLine(
                FormatException(e.Exception))

            AppendRecord(
                "UNOBSERVED_TASK_EXCEPTION",
                details.ToString())

            'Do not call e.SetObserved() here.
            'The logger should not change application behaviour.

        End Sub

        Private Shared Function CurrentDomain_AssemblyResolve(
            ByVal sender As System.Object,
            ByVal e As System.ResolveEventArgs) As System.Reflection.Assembly

            Dim details As New System.Text.StringBuilder()

            details.AppendLine(
                "RequestedAssembly=" &
                If(e.Name, System.String.Empty))

            If e.RequestingAssembly IsNot Nothing Then

                details.AppendLine(
                    "RequestingAssembly=" &
                    e.RequestingAssembly.FullName)

                details.AppendLine(
                    "RequestingAssemblyLocation=" &
                    GetAssemblyLocation(e.RequestingAssembly))

            End If

            AppendRecord(
                "ASSEMBLY_RESOLVE_FAILED",
                details.ToString())

            'Never load an arbitrary DLL from this diagnostic handler.
            Return Nothing

        End Function

        Private Shared Sub CurrentDomain_ProcessExit(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs)

            AppendRecord(
                "PROCESS_EXIT",
                "The Office process is exiting normally.")

        End Sub

        Private Shared Sub CurrentDomain_DomainUnload(
            ByVal sender As System.Object,
            ByVal e As System.EventArgs)

            AppendRecord(
                "APPDOMAIN_UNLOAD",
                "The add-in AppDomain is being unloaded.")

        End Sub

        Private Shared Sub DeferredSnapshotTimerCallback(
            ByVal state As System.Object)

            Try

                If DiagnosticsEnabled Then

                    AppendRecord(
                        "EXTENDED_SNAPSHOT",
                        BuildExtendedSnapshot())

                End If

            Catch ex As System.Exception

                AppendRecord(
                    "EXTENDED_SNAPSHOT_ERROR",
                    FormatException(ex))

            Finally

                Dim timerToDispose As System.Threading.Timer =
                    Nothing

                SyncLock InitializationLock

                    timerToDispose = SnapshotTimer
                    SnapshotTimer = Nothing

                End SyncLock

                If timerToDispose IsNot Nothing Then

                    Try
                        timerToDispose.Dispose()
                    Catch ex As System.Exception
                    End Try

                End If

            End Try

        End Sub

        Private Shared Function BuildStartupInformation() As System.String

            Dim result As New System.Text.StringBuilder()

            result.AppendLine(
                "LocalTime=" &
                System.DateTime.Now.ToString(
                    "yyyy-MM-dd HH:mm:ss.fff zzz",
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "UtcTime=" &
                System.DateTime.UtcNow.ToString(
                    "o",
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "MachineName=" &
                System.Environment.MachineName)

            result.AppendLine(
                "OSVersion=" &
                System.Environment.OSVersion.VersionString)

            result.AppendLine(
                "CLRVersion=" &
                System.Environment.Version.ToString())

            result.AppendLine(
                "Is64BitOperatingSystem=" &
                System.Environment.Is64BitOperatingSystem.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "Is64BitProcess=" &
                System.Environment.Is64BitProcess.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "ProcessorCount=" &
                System.Environment.ProcessorCount.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "CurrentCulture=" &
                System.Globalization.CultureInfo.CurrentCulture.Name)

            result.AppendLine(
                "CurrentUICulture=" &
                System.Globalization.CultureInfo.CurrentUICulture.Name)

            result.AppendLine(
                "InstalledUICulture=" &
                System.Globalization.CultureInfo.InstalledUICulture.Name)

            result.AppendLine(
                "UserInteractive=" &
                System.Environment.UserInteractive.ToString(
                    System.Globalization.CultureInfo.InvariantCulture))

            result.AppendLine(
                "AppDomainName=" &
                System.AppDomain.CurrentDomain.FriendlyName)

            result.AppendLine(
                "AppDomainBaseDirectory=" &
                SafeGetAppDomainBaseDirectory())

            AppendSelectedEnvironmentVariable(
                result,
                "VSTO_LOGALERTS")

            AppendSelectedEnvironmentVariable(
                result,
                "VSTO_SUPPRESSDISPLAYALERTS")

            AppendSelectedEnvironmentVariable(
                result,
                "COR_ENABLE_PROFILING")

            AppendSelectedEnvironmentVariable(
                result,
                "COR_PROFILER")

            AppendSelectedEnvironmentVariable(
                result,
                "CORECLR_ENABLE_PROFILING")

            AppendSelectedEnvironmentVariable(
                result,
                "CORECLR_PROFILER")

            AppendSelectedEnvironmentVariable(
                result,
                "__COMPAT_LAYER")

            Try

                Using currentProcess As System.Diagnostics.Process =
                    System.Diagnostics.Process.GetCurrentProcess()

                    result.AppendLine(
                        "ProcessName=" &
                        currentProcess.ProcessName)

                    result.AppendLine(
                        "ProcessId=" &
                        currentProcess.Id.ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

                    result.AppendLine(
                        "ProcessStartTimeUtc=" &
                        currentProcess.StartTime.
                            ToUniversalTime().
                            ToString(
                                "o",
                                System.Globalization.CultureInfo.InvariantCulture))

                    result.AppendLine(
                        "WorkingSetBytes=" &
                        currentProcess.WorkingSet64.ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

                    result.AppendLine(
                        "PrivateMemoryBytes=" &
                        currentProcess.PrivateMemorySize64.ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

                    Try

                        result.AppendLine(
                            "HostExecutable=" &
                            currentProcess.MainModule.FileName)

                        result.AppendLine(
                            "HostFileVersion=" &
                            currentProcess.MainModule.
                                FileVersionInfo.
                                FileVersion)

                        result.AppendLine(
                            "HostProductVersion=" &
                            currentProcess.MainModule.
                                FileVersionInfo.
                                ProductVersion)

                    Catch ex As System.Exception

                        result.AppendLine(
                            "HostExecutableInformationError=" &
                            ex.Message)

                    End Try

                End Using

            Catch ex As System.Exception

                result.AppendLine(
                    "ProcessInformationError=" &
                    ex.ToString())

            End Try

            AppendAddInAssemblyInformation(result)

            Return result.ToString()

        End Function

        Private Shared Sub AppendAddInAssemblyInformation(
            ByVal result As System.Text.StringBuilder)

            If AddInAssembly Is Nothing Then

                result.AppendLine(
                    "AddInAssembly=Not supplied")

                Return

            End If

            Try

                result.AppendLine(
                    "AddInAssemblyFullName=" &
                    AddInAssembly.FullName)

                result.AppendLine(
                    "AddInAssemblyRuntimeVersion=" &
                    AddInAssembly.ImageRuntimeVersion)

                Dim assemblyLocation As System.String =
                    GetAssemblyLocation(AddInAssembly)

                result.AppendLine(
                    "AddInAssemblyLocation=" &
                    assemblyLocation)

                If Not System.String.IsNullOrWhiteSpace(
                    assemblyLocation) AndAlso
                   System.IO.File.Exists(assemblyLocation) Then

                    Dim fileInformation As New System.IO.FileInfo(
                        assemblyLocation)

                    result.AppendLine(
                        "AddInAssemblySize=" &
                        fileInformation.Length.ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

                    result.AppendLine(
                        "AddInAssemblyModifiedUtc=" &
                        fileInformation.LastWriteTimeUtc.ToString(
                            "o",
                            System.Globalization.CultureInfo.InvariantCulture))

                    result.AppendLine(
                        "AddInAssemblySHA256=" &
                        CalculateSha256(assemblyLocation))

                    Try

                        Dim versionInformation As System.Diagnostics.FileVersionInfo =
                            System.Diagnostics.FileVersionInfo.GetVersionInfo(
                                assemblyLocation)

                        result.AppendLine(
                            "AddInAssemblyFileVersion=" &
                            versionInformation.FileVersion)

                        result.AppendLine(
                            "AddInAssemblyProductVersion=" &
                            versionInformation.ProductVersion)

                    Catch ex As System.Exception

                        result.AppendLine(
                            "AddInAssemblyVersionError=" &
                            ex.Message)

                    End Try

                End If

            Catch ex As System.Exception

                result.AppendLine(
                    "AddInAssemblyInformationError=" &
                    ex.ToString())

            End Try

        End Sub

        Private Shared Function BuildExtendedSnapshot() As System.String

            Dim result As New System.Text.StringBuilder()

            result.AppendLine(
                "SnapshotUtc=" &
                System.DateTime.UtcNow.ToString(
                    "o",
                    System.Globalization.CultureInfo.InvariantCulture))

            AppendManagedAssemblies(result)
            AppendProcessModules(result)
            AppendAddInDirectoryFiles(result)
            AppendRegisteredOfficeAddIns(result)

            Return result.ToString()

        End Function

        Private Shared Sub AppendManagedAssemblies(
            ByVal result As System.Text.StringBuilder)

            result.AppendLine()
            result.AppendLine("----- LOADED MANAGED ASSEMBLIES -----")

            Try

                Dim assemblyLines As New System.Collections.Generic.List(
                    Of System.String)()

                For Each loadedAssembly As System.Reflection.Assembly In
                    System.AppDomain.CurrentDomain.GetAssemblies()

                    Dim line As System.String

                    Try

                        line =
                            loadedAssembly.FullName &
                            " | Dynamic=" &
                            loadedAssembly.IsDynamic.ToString(
                                System.Globalization.CultureInfo.InvariantCulture) &
                            " | Location=" &
                            GetAssemblyLocation(loadedAssembly)

                    Catch ex As System.Exception

                        line =
                            "AssemblyInformationError=" &
                            ex.Message

                    End Try

                    assemblyLines.Add(line)

                Next

                assemblyLines.Sort(
                    System.StringComparer.OrdinalIgnoreCase)

                For Each line As System.String In assemblyLines
                    result.AppendLine(line)
                Next

            Catch ex As System.Exception

                result.AppendLine(
                    "ManagedAssemblySnapshotError=" &
                    ex.ToString())

            End Try

        End Sub

        Private Shared Sub AppendProcessModules(
            ByVal result As System.Text.StringBuilder)

            result.AppendLine()
            result.AppendLine("----- LOADED PROCESS MODULES -----")

            Try

                Dim moduleLines As New System.Collections.Generic.List(
                    Of System.String)()

                Using currentProcess As System.Diagnostics.Process =
                    System.Diagnostics.Process.GetCurrentProcess()

                    For Each processModule As System.Diagnostics.ProcessModule In
                        currentProcess.Modules

                        Try

                            Dim modulePath As System.String =
                                processModule.FileName

                            Dim moduleVersion As System.String =
                                processModule.FileVersionInfo.FileVersion

                            Dim moduleSize As System.Int64 = 0

                            Try

                                If System.IO.File.Exists(modulePath) Then

                                    moduleSize =
                                        New System.IO.FileInfo(
                                            modulePath).Length

                                End If

                            Catch ex As System.Exception
                            End Try

                            moduleLines.Add(
                                processModule.ModuleName &
                                " | Version=" &
                                moduleVersion &
                                " | Size=" &
                                moduleSize.ToString(
                                    System.Globalization.CultureInfo.InvariantCulture) &
                                " | Path=" &
                                modulePath)

                        Catch ex As System.Exception

                            moduleLines.Add(
                                "ModuleInformationError=" &
                                ex.Message)

                        End Try

                    Next

                End Using

                moduleLines.Sort(
                    System.StringComparer.OrdinalIgnoreCase)

                For Each line As System.String In moduleLines
                    result.AppendLine(line)
                Next

            Catch ex As System.Exception

                result.AppendLine(
                    "ProcessModuleSnapshotError=" &
                    ex.ToString())

            End Try

        End Sub

        Private Shared Sub AppendAddInDirectoryFiles(
            ByVal result As System.Text.StringBuilder)

            result.AppendLine()
            result.AppendLine("----- ADD-IN DIRECTORY FILES -----")

            Dim assemblyLocation As System.String =
                GetAssemblyLocation(AddInAssembly)

            If System.String.IsNullOrWhiteSpace(
                assemblyLocation) Then

                result.AppendLine(
                    "Add-in assembly location is unavailable.")

                Return

            End If

            Dim directoryPath As System.String =
                System.IO.Path.GetDirectoryName(
                    assemblyLocation)

            If System.String.IsNullOrWhiteSpace(
                directoryPath) OrElse
               Not System.IO.Directory.Exists(directoryPath) Then

                result.AppendLine(
                    "Add-in directory does not exist.")

                Return

            End If

            Try

                Dim fileNames() As System.String =
                    System.IO.Directory.GetFiles(
                        directoryPath,
                        "*",
                        System.IO.SearchOption.TopDirectoryOnly)

                System.Array.Sort(
                    fileNames,
                    System.StringComparer.OrdinalIgnoreCase)

                Dim writtenFileCount As System.Int32 = 0

                For Each fileName As System.String In fileNames

                    If writtenFileCount >= 250 Then

                        result.AppendLine(
                            "File list truncated after 250 files.")

                        Exit For

                    End If

                    Try

                        Dim fileInformation As New System.IO.FileInfo(
                            fileName)

                        Dim fileVersion As System.String =
                            System.String.Empty

                        Try

                            fileVersion =
                                System.Diagnostics.FileVersionInfo.
                                    GetVersionInfo(fileName).
                                    FileVersion

                        Catch ex As System.Exception
                        End Try

                        result.AppendLine(
                            fileInformation.Name &
                            " | Version=" &
                            fileVersion &
                            " | Size=" &
                            fileInformation.Length.ToString(
                                System.Globalization.CultureInfo.InvariantCulture) &
                            " | ModifiedUtc=" &
                            fileInformation.LastWriteTimeUtc.ToString(
                                "o",
                                System.Globalization.CultureInfo.InvariantCulture))

                        writtenFileCount += 1

                    Catch ex As System.Exception

                        result.AppendLine(
                            "FileInformationError=" &
                            ex.Message)

                    End Try

                Next

            Catch ex As System.Exception

                result.AppendLine(
                    "AddInDirectorySnapshotError=" &
                    ex.ToString())

            End Try

        End Sub

        Private Shared Sub AppendRegisteredOfficeAddIns(
            ByVal result As System.Text.StringBuilder)

            result.AppendLine()
            result.AppendLine("----- REGISTERED OFFICE ADD-INS -----")

            Dim officeApplicationName As System.String =
                GetOfficeApplicationName()

            If System.String.IsNullOrWhiteSpace(
                officeApplicationName) Then

                result.AppendLine(
                    "The current process is not a recognized Office host.")

                Return

            End If

            Dim registryPath As System.String =
                "Software\Microsoft\Office\" &
                officeApplicationName &
                "\Addins"

            Dim registryViews As New System.Collections.Generic.List(
                Of Microsoft.Win32.RegistryView)()

            registryViews.Add(
                Microsoft.Win32.RegistryView.Registry32)

            If System.Environment.Is64BitOperatingSystem Then

                registryViews.Add(
                    Microsoft.Win32.RegistryView.Registry64)

            End If

            Dim registryHives() As Microsoft.Win32.RegistryHive = {
                Microsoft.Win32.RegistryHive.CurrentUser,
                Microsoft.Win32.RegistryHive.LocalMachine
            }

            For Each registryHive As Microsoft.Win32.RegistryHive In
                registryHives

                For Each registryView As Microsoft.Win32.RegistryView In
                    registryViews

                    AppendRegisteredOfficeAddInsForRegistryLocation(
                        result,
                        registryHive,
                        registryView,
                        registryPath)

                Next

            Next

        End Sub

        Private Shared Sub AppendRegisteredOfficeAddInsForRegistryLocation(
            ByVal result As System.Text.StringBuilder,
            ByVal registryHive As Microsoft.Win32.RegistryHive,
            ByVal registryView As Microsoft.Win32.RegistryView,
            ByVal registryPath As System.String)

            Try

                Using baseKey As Microsoft.Win32.RegistryKey =
                    Microsoft.Win32.RegistryKey.OpenBaseKey(
                        registryHive,
                        registryView)

                    Using addInsKey As Microsoft.Win32.RegistryKey =
                        baseKey.OpenSubKey(
                            registryPath,
                            False)

                        If addInsKey Is Nothing Then
                            Return
                        End If

                        Dim addInSubKeyNames() As System.String =
                            addInsKey.GetSubKeyNames()

                        System.Array.Sort(
                            addInSubKeyNames,
                            System.StringComparer.OrdinalIgnoreCase)

                        For Each addInSubKeyName As System.String In
                            addInSubKeyNames

                            Try

                                Using addInKey As Microsoft.Win32.RegistryKey =
                                    addInsKey.OpenSubKey(
                                        addInSubKeyName,
                                        False)

                                    If addInKey Is Nothing Then
                                        Continue For
                                    End If

                                    result.AppendLine(
                                        registryHive.ToString() &
                                        "\" &
                                        registryView.ToString() &
                                        " | ProgId=" &
                                        addInSubKeyName &
                                        " | LoadBehavior=" &
                                        RegistryValueToString(
                                            addInKey.GetValue(
                                                "LoadBehavior",
                                                Nothing)) &
                                        " | FriendlyName=" &
                                        RegistryValueToString(
                                            addInKey.GetValue(
                                                "FriendlyName",
                                                Nothing)) &
                                        " | Description=" &
                                        RegistryValueToString(
                                            addInKey.GetValue(
                                                "Description",
                                                Nothing)) &
                                        " | Manifest=" &
                                        RegistryValueToString(
                                            addInKey.GetValue(
                                                "Manifest",
                                                Nothing)))

                                End Using

                            Catch ex As System.Exception

                                result.AppendLine(
                                    registryHive.ToString() &
                                    "\" &
                                    registryView.ToString() &
                                    " | ProgId=" &
                                    addInSubKeyName &
                                    " | Error=" &
                                    ex.Message)

                            End Try

                        Next

                    End Using

                End Using

            Catch ex As System.Exception

                result.AppendLine(
                    registryHive.ToString() &
                    "\" &
                    registryView.ToString() &
                    " | RegistryReadError=" &
                    ex.Message)

            End Try

        End Sub

        Private Shared Function GetOfficeApplicationName() As System.String

            Try

                Using currentProcess As System.Diagnostics.Process =
                    System.Diagnostics.Process.GetCurrentProcess()

                    Select Case currentProcess.ProcessName.ToUpperInvariant()

                        Case "WINWORD"
                            Return "Word"

                        Case "OUTLOOK"
                            Return "Outlook"

                        Case "EXCEL"
                            Return "Excel"

                        Case "POWERPNT"
                            Return "PowerPoint"

                        Case Else
                            Return System.String.Empty

                    End Select

                End Using

            Catch ex As System.Exception
                Return System.String.Empty
            End Try

        End Function

        Private Shared Function RegistryValueToString(
            ByVal value As System.Object) As System.String

            If value Is Nothing Then
                Return System.String.Empty
            End If

            Try

                Dim byteArray() As System.Byte =
                    TryCast(
                        value,
                        System.Byte())

                If byteArray IsNot Nothing Then

                    Return "ByteArray[" &
                        byteArray.Length.ToString(
                            System.Globalization.CultureInfo.InvariantCulture) &
                        "]"

                End If

                Return SafeSingleLine(
                    System.Convert.ToString(
                        value,
                        System.Globalization.CultureInfo.InvariantCulture))

            Catch ex As System.Exception
                Return "ValueReadError=" & ex.Message
            End Try

        End Function

        Private Shared Sub AppendSelectedEnvironmentVariable(
            ByVal result As System.Text.StringBuilder,
            ByVal variableName As System.String)

            Try

                Dim value As System.String =
                    System.Environment.GetEnvironmentVariable(
                        variableName)

                If Not System.String.IsNullOrWhiteSpace(value) Then

                    result.AppendLine(
                        variableName &
                        "=" &
                        SafeSingleLine(value))

                End If

            Catch ex As System.Exception

                result.AppendLine(
                    variableName &
                    "=ReadError:" &
                    ex.Message)

            End Try

        End Sub

        Private Shared Function SafeGetAppDomainBaseDirectory() As System.String

            Try
                Return System.AppDomain.CurrentDomain.BaseDirectory
            Catch ex As System.Exception
                Return "Unavailable: " & ex.Message
            End Try

        End Function

        Private Shared Function GetAssemblyLocation(
            ByVal assembly As System.Reflection.Assembly) As System.String

            If assembly Is Nothing Then
                Return System.String.Empty
            End If

            Try

                If assembly.IsDynamic Then
                    Return "[Dynamic assembly]"
                End If

                Return assembly.Location

            Catch ex As System.Exception
                Return "[Location unavailable: " & ex.Message & "]"
            End Try

        End Function

        Private Shared Function CalculateSha256(
            ByVal filePath As System.String) As System.String

            Try

                Using fileStream As New System.IO.FileStream(
                    filePath,
                    System.IO.FileMode.Open,
                    System.IO.FileAccess.Read,
                    System.IO.FileShare.ReadWrite Or
                        System.IO.FileShare.Delete)

                    Using sha256 As System.Security.Cryptography.SHA256 =
                        System.Security.Cryptography.SHA256.Create()

                        Dim hashBytes() As System.Byte =
                            sha256.ComputeHash(fileStream)

                        Dim result As New System.Text.StringBuilder(
                            hashBytes.Length * 2)

                        For Each hashByte As System.Byte In hashBytes

                            result.Append(
                                hashByte.ToString(
                                    "x2",
                                    System.Globalization.CultureInfo.InvariantCulture))

                        Next

                        Return result.ToString()

                    End Using

                End Using

            Catch ex As System.Exception
                Return "SHA256Error=" & ex.Message
            End Try

        End Function

        Private Shared Function FormatException(
            ByVal ex As System.Exception) As System.String

            If ex Is Nothing Then
                Return "No exception object was supplied."
            End If

            Dim result As New System.Text.StringBuilder()

            Try

                result.AppendLine(
                    "ExceptionType=" &
                    ex.GetType().FullName)

                result.AppendLine(
                    "HResult=0x" &
                    ex.HResult.ToString(
                        "X8",
                        System.Globalization.CultureInfo.InvariantCulture))

                result.AppendLine(
                    "Message=" &
                    ex.Message)

                result.AppendLine(
                    "Source=" &
                    If(ex.Source, System.String.Empty))

                If ex.TargetSite IsNot Nothing Then

                    result.AppendLine(
                        "TargetSite=" &
                        ex.TargetSite.ToString())

                End If

                result.AppendLine()
                result.AppendLine(ex.ToString())

                Dim reflectionTypeLoadException As System.Reflection.ReflectionTypeLoadException = TryCast(ex, System.Reflection.ReflectionTypeLoadException)

                If reflectionTypeLoadException IsNot Nothing AndAlso
                   reflectionTypeLoadException.LoaderExceptions IsNot Nothing Then

                    result.AppendLine()
                    result.AppendLine("Loader exceptions:")

                    For Each loaderException As System.Exception In
                        reflectionTypeLoadException.LoaderExceptions

                        If loaderException IsNot Nothing Then

                            result.AppendLine(
                                loaderException.ToString())

                        End If

                    Next

                End If

                Dim fileLoadException As System.IO.FileLoadException =
                    TryCast(
                        ex,
                        System.IO.FileLoadException)

                If fileLoadException IsNot Nothing Then

                    Try

                        If Not System.String.IsNullOrWhiteSpace(
                            fileLoadException.FusionLog) Then

                            result.AppendLine()
                            result.AppendLine("Fusion log:")
                            result.AppendLine(
                                fileLoadException.FusionLog)

                        End If

                    Catch fusionException As System.Exception

                        result.AppendLine(
                            "FusionLogReadError=" &
                            fusionException.Message)

                    End Try

                End If

                If ex.Data IsNot Nothing AndAlso
                   ex.Data.Count > 0 Then

                    result.AppendLine()
                    result.AppendLine("Exception data:")

                    For Each item As System.Collections.DictionaryEntry In
                        ex.Data

                        result.AppendLine(
                            SafeToString(item.Key) &
                            "=" &
                            SafeToString(item.Value))

                    Next

                End If

            Catch formattingException As System.Exception

                result.AppendLine(
                    "ExceptionFormattingError=" &
                    formattingException.Message)

            End Try

            Return result.ToString()

        End Function

        Private Shared Function SafeToString(
            ByVal value As System.Object) As System.String

            If value Is Nothing Then
                Return System.String.Empty
            End If

            Try
                Return value.ToString()
            Catch ex As System.Exception
                Return "[ToString failed: " & ex.Message & "]"
            End Try

        End Function

        Private Shared Function SafeSingleLine(
            ByVal value As System.String) As System.String

            If value Is Nothing Then
                Return System.String.Empty
            End If

            Return value.
                Replace(
                    System.Convert.ToString(
                        System.Convert.ToChar(13)),
                    " ").
                Replace(
                    System.Convert.ToString(
                        System.Convert.ToChar(10)),
                    " ").
                Replace(
                    System.Convert.ToString(
                        System.Convert.ToChar(9)),
                    " ")

        End Function

        Private Shared Sub AppendRecord(
            ByVal eventName As System.String,
            ByVal details As System.String)

            If Not DiagnosticsEnabled Then
                Return
            End If

            If System.String.IsNullOrWhiteSpace(
                LogFilePathValue) Then

                Return

            End If

            Dim record As New System.Text.StringBuilder()

            record.AppendLine()
            record.AppendLine(
                "========================================================================")

            record.AppendLine(
                "TimestampUtc=" &
                System.DateTime.UtcNow.ToString(
                    "o",
                    System.Globalization.CultureInfo.InvariantCulture))

            record.AppendLine(
                "TimestampLocal=" &
                System.DateTime.Now.ToString(
                    "yyyy-MM-dd HH:mm:ss.fff zzz",
                    System.Globalization.CultureInfo.InvariantCulture))

            record.AppendLine(
                "SessionId=" &
                SessionId)

            record.AppendLine(
                "AddIn=" &
                AddInName)

            If Not System.String.IsNullOrWhiteSpace(
                RdvVersionValue) Then

                record.AppendLine(
                    "RDV=" &
                    RdvVersionValue)

            End If

            record.AppendLine(
                "Event=" &
                If(eventName, System.String.Empty))

            Try

                record.AppendLine(
                    "ProcessId=" &
                    System.Diagnostics.Process.
                        GetCurrentProcess().
                        Id.
                        ToString(
                            System.Globalization.CultureInfo.InvariantCulture))

            Catch ex As System.Exception
            End Try

            record.AppendLine(
                "ManagedThreadId=" &
                System.Threading.Thread.
                    CurrentThread.
                    ManagedThreadId.
                    ToString(
                        System.Globalization.CultureInfo.InvariantCulture))

            Try

                record.AppendLine(
                    "ThreadApartmentState=" &
                    System.Threading.Thread.
                        CurrentThread.
                        GetApartmentState().
                        ToString())

            Catch ex As System.Exception
            End Try

            If Not System.String.IsNullOrWhiteSpace(details) Then

                record.AppendLine()
                record.AppendLine(details)

            End If

            WriteTextToLogFile(
                record.ToString())

        End Sub

        Private Shared Sub WriteTextToLogFile(
            ByVal text As System.String)

            Dim fileMutex As System.Threading.Mutex =
                Nothing

            Dim mutexOwned As System.Boolean =
                False

            Try

                fileMutex =
                    New System.Threading.Mutex(
                        False,
                        FileMutexName)

                Try

                    mutexOwned =
                        fileMutex.WaitOne(500)

                Catch ex As System.Threading.AbandonedMutexException

                    'An abandoned mutex means this process now owns it.
                    mutexOwned = True

                Catch ex As System.Exception

                    mutexOwned = False

                End Try

                Try

                    System.IO.Directory.CreateDirectory(
                        System.IO.Path.GetDirectoryName(
                            LogFilePathValue))

                Catch ex As System.Exception
                    Return
                End Try

                If mutexOwned Then
                    RotateLogFileIfRequired()
                End If

                Using fileStream As New System.IO.FileStream(
                    LogFilePathValue,
                    System.IO.FileMode.Append,
                    System.IO.FileAccess.Write,
                    System.IO.FileShare.ReadWrite Or
                        System.IO.FileShare.Delete)

                    Using writer As New System.IO.StreamWriter(
                        fileStream,
                        New System.Text.UTF8Encoding(False))

                        writer.Write(text)
                        writer.Flush()

                    End Using

                End Using

            Catch ex As System.Exception

                'The logger must never interfere with Word or Outlook.

            Finally

                If mutexOwned AndAlso
                   fileMutex IsNot Nothing Then

                    Try
                        fileMutex.ReleaseMutex()
                    Catch ex As System.Exception
                    End Try

                End If

                If fileMutex IsNot Nothing Then

                    Try
                        fileMutex.Dispose()
                    Catch ex As System.Exception
                    End Try

                End If

            End Try

        End Sub

        Private Shared Sub RotateLogFileIfRequired()

            Try

                If Not System.IO.File.Exists(
                    LogFilePathValue) Then

                    Return

                End If

                Dim fileInformation As New System.IO.FileInfo(
                    LogFilePathValue)

                If fileInformation.Length <
                   MaximumLogFileSize Then

                    Return

                End If

                Dim previousLogPath As System.String =
                    System.IO.Path.Combine(
                        fileInformation.DirectoryName,
                        "RI-CrashLog.previous.txt")

                If System.IO.File.Exists(
                    previousLogPath) Then

                    Try
                        System.IO.File.Delete(
                            previousLogPath)
                    Catch ex As System.Exception
                        Return
                    End Try

                End If

                System.IO.File.Move(
                    LogFilePathValue,
                    previousLogPath)

            Catch ex As System.Exception
            End Try

        End Sub

        Private Shared Sub WriteCollectorReadme(
    ByVal targetDirectory As System.String)

            Try

                Dim readmePath As System.String =
                    System.IO.Path.Combine(
                        targetDirectory,
                        "RI-CrashLog-README.txt")

                Dim content As New System.Text.StringBuilder()

                content.AppendLine("Red Ink - Crash Diagnostics Collection Guide")
                content.AppendLine("============================================")
                content.AppendLine()
                content.AppendLine("This machine has Red Ink crash diagnostics enabled for a limited time.")
                content.AppendLine("The add-in writes stability information to 'RI_CrashLog.txt' (and, once it")
                content.AppendLine("reaches about 10 MB, 'RI_CrashLog.previous.txt') in this same folder.")
                content.AppendLine("If DLL diagnostics are generated, 'RI_DLL_Loaded.txt' is written to this")
                content.AppendLine("same folder as well.")
                content.AppendLine()
                content.AppendLine("WHAT TO COLLECT WHEN A HOST (WORD / OUTLOOK / EXCEL) CRASHES OR HANGS")
                content.AppendLine("--------------------------------------------------------------------")
                content.AppendLine("Please gather the following and send them to the Red Ink support contact,")
                content.AppendLine("together with the approximate date and time of the problem:")
                content.AppendLine()
                content.AppendLine("1. From this folder:")
                content.AppendLine("   " & targetDirectory)
                content.AppendLine("   Collect these files only:")
                content.AppendLine("   - RI_CrashLog.txt")
                content.AppendLine("   - RI_CrashLog.previous.txt (if available)")
                content.AppendLine("   - RI_DLL_Loaded.txt (if available)")
                content.AppendLine()
                content.AppendLine("2. Windows Event Log entries around the time of the crash:")
                content.AppendLine("   - Open 'Event Viewer' (eventvwr.msc).")
                content.AppendLine("   - Go to 'Windows Logs' > 'Application'.")
                content.AppendLine("   - Look for 'Error' entries with source 'Application Error',")
                content.AppendLine("     '.NET Runtime', 'VSTO 4.0', 'Office' or the host name")
                content.AppendLine("     (WINWORD.EXE, OUTLOOK.EXE, EXCEL.EXE).")
                content.AppendLine("   - Right-click each relevant entry > 'Copy' > 'Copy Details as Text',")
                content.AppendLine("     or select them and use 'Save Selected Events...' (.evtx).")
                content.AppendLine()
                content.AppendLine("3. The list of Office 'Disabled Items' and COM add-ins for the affected host:")
                content.AppendLine("   - In the host: File > Options > Add-ins.")
                content.AppendLine("   - At the bottom, in 'Manage', check both 'COM Add-ins' and")
                content.AppendLine("     'Disabled Items', and note whether 'Red Ink' appears as disabled.")
                content.AppendLine("   - A screenshot of each dialog is sufficient.")
                content.AppendLine()
                content.AppendLine("4. Office and Windows version information:")
                content.AppendLine("   - In the host: File > Account > 'About <host>' (note the exact build number).")
                content.AppendLine("   - Windows: run 'winver' and note the version/build.")
                content.AppendLine()
                content.AppendLine("TURNING DIAGNOSTICS OFF")
                content.AppendLine("-----------------------")
                content.AppendLine("Diagnostics are controlled by the 'Crashlog' setting in the Red Ink")
                content.AppendLine("configuration (redink.ini). Setting it back to off (0) disables collection")
                content.AppendLine("on the next start of the host application.")

                System.IO.File.WriteAllText(
                    readmePath,
                    content.ToString(),
                    New System.Text.UTF8Encoding(False))

            Catch ex As System.Exception

                'The logger must never interfere with the host application.

            End Try

        End Sub



    End Class

End Namespace
