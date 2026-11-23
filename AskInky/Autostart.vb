' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: Autostart.vb
' Purpose: Start Ask Inky automatically at Windows sign-in, per user, as set in the settings.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Part of Red Ink (AskInky). P18 W367 (DARO UID 569): one value "Ask Inky" in HKCU\...\Run, no admin rights.
' The installer writes the same value for a per-user installation and removes it on uninstall; Ask Inky
' keeps it in line with the setting "Mit Windows starten" (default on) and with its current location.
Friend NotInheritable Class Autostart
    Private Const RunKey As String = "Software\Microsoft\Windows\CurrentVersion\Run"
    Private Const ValueName As String = "Ask Inky"
    Private Sub New()
    End Sub

    ' Default for a new installation: an admin policy (GPO/Intune, StartWithWindowsDefault under Software\Policies),
    ' else the installer's choice (MSI property AUTOSTART, written by the per-user MSI under HKCU\Software\Red Ink\AskInky),
    ' otherwise on. 2.0.2 (DARO UID 571): the installer no longer writes under Software\Policies.
    Friend Shared Function InstallerDefault() As Boolean
        For Each location As System.Tuple(Of Microsoft.Win32.RegistryKey, String) In {
                System.Tuple.Create(Microsoft.Win32.Registry.LocalMachine, "Software\Policies\Red Ink\AskInky"),
                System.Tuple.Create(Microsoft.Win32.Registry.CurrentUser, "Software\Policies\Red Ink\AskInky"),
                System.Tuple.Create(Microsoft.Win32.Registry.CurrentUser, "Software\Red Ink\AskInky")}
            Try
                Using key As Microsoft.Win32.RegistryKey = location.Item1.OpenSubKey(location.Item2, False)
                    If key Is Nothing Then Continue For
                    Dim value As Object = key.GetValue("StartWithWindowsDefault")
                    If value IsNot Nothing Then Return Not String.Equals(System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture).Trim(), "0", System.StringComparison.Ordinal)
                End Using
            Catch ex As System.Exception
                ClientLog.Write("autostart.default.read.failed", ex)
            End Try
        Next
        Return True
    End Function

    Friend Shared Function IsEnabled() As Boolean
        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunKey, False)
                Return key IsNot Nothing AndAlso key.GetValue(ValueName) IsNot Nothing
            End Using
        Catch
            Return False
        End Try
    End Function

    ' Writes or removes the per-user Run value. Returns False when the registry refused it.
    Friend Shared Function Apply(enabled As Boolean) As Boolean
        Try
            Using key As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunKey, True)
                If enabled Then
                    Dim command As String = """" & System.Environment.ProcessPath & """ " & Program.AutostartArgument
                    If Not String.Equals(System.Convert.ToString(key.GetValue(ValueName), System.Globalization.CultureInfo.InvariantCulture), command, System.StringComparison.OrdinalIgnoreCase) Then
                        key.SetValue(ValueName, command, Microsoft.Win32.RegistryValueKind.String)
                        ClientLog.Write("autostart.enabled")
                    End If
                ElseIf key.GetValue(ValueName) IsNot Nothing Then
                    key.DeleteValue(ValueName, False)
                    ClientLog.Write("autostart.disabled")
                End If
            End Using
            Return True
        Catch ex As System.Exception
            ClientLog.Write("autostart.apply.failed", ex)
            Return False
        End Try
    End Function
End Class
