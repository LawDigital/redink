' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.Security.vb
' Purpose:
'   Host-specific browser-extension authorization for the Word localhost receiver,
'   using the shared origin allow-list and Red Ink custom authorization dialog.
'
' Regression invariants:
'   - The existing port, JSON payload and redink_sendtoword command remain unchanged.
'   - The installed Chromium extension needs no protocol/token change; its origin is
'     approved once by the user and then remembered where My.Settings is available.
'   - Missing My.Settings entries retain approvals only in memory for the current
'     Word process. Missing persistence never falls back to wildcard access.
' =============================================================================

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn

    Private Const LocalHttpAuthMasterKeySettingName As System.String = "LocalHttpAuthMasterKey"
    Private Const LocalHttpApprovedOriginsSettingName As System.String = "LocalHttpApprovedOrigins"

    Private ReadOnly _localHttpSecurityGate As New System.Object()
    Private ReadOnly _localHttpApprovalGate As New System.Threading.SemaphoreSlim(1, 1)
    Private _localHttpSecurity As SharedLibrary.SharedLibrary.LocalHttpBrowserSecurity = Nothing

    Private Function GetLocalHttpSecurity() As SharedLibrary.SharedLibrary.LocalHttpBrowserSecurity
        SyncLock _localHttpSecurityGate
            If _localHttpSecurity Is Nothing Then
                _localHttpSecurity = New SharedLibrary.SharedLibrary.LocalHttpBrowserSecurity(
                    LocalHttpAuthMasterKeySettingName,
                    LocalHttpApprovedOriginsSettingName,
                    AddressOf ReadLocalHttpSecuritySetting,
                    AddressOf WriteLocalHttpSecuritySetting,
                    AddressOf LogLocalHttpSecurity)
            End If

            Return _localHttpSecurity
        End SyncLock
    End Function

    Private Sub LogLocalHttpSecurity(ByVal message As System.String)
        System.Diagnostics.Debug.WriteLine("[LocalHttpSecurity] " & If(message, System.String.Empty))
    End Sub

    Private Function ReadLocalHttpSecuritySetting(ByVal settingName As System.String) As System.String
        Try
            Dim value As System.Object = My.Settings(settingName)
            If value Is Nothing Then Return System.String.Empty
            Return System.Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture)
        Catch ex As System.Exception
            LogLocalHttpSecurity("Cannot read My.Settings entry '" & settingName & "' (" & ex.GetType().Name & ").")
            Return System.String.Empty
        End Try
    End Function

    Private Function WriteLocalHttpSecuritySetting(
        ByVal settingName As System.String,
        ByVal settingValue As System.String
    ) As System.Boolean
        Try
            My.Settings(settingName) = If(settingValue, System.String.Empty)
            My.Settings.Save()
            Return True
        Catch ex As System.Exception
            LogLocalHttpSecurity("Cannot write My.Settings entry '" & settingName & "' (" & ex.GetType().Name & ").")
            Return False
        End Try
    End Function

    Private Function IsLocalLoopbackRequest(ByVal req As System.Net.HttpListenerRequest) As System.Boolean
        Return GetLocalHttpSecurity().IsLoopbackRequest(req)
    End Function

    Private Async Function EnsureApprovedBrowserExtensionOriginAsync(
        ByVal req As System.Net.HttpListenerRequest
    ) As System.Threading.Tasks.Task(Of System.String)
        Dim origin As System.String = If(req.Headers("Origin"), System.String.Empty)
        Dim security As SharedLibrary.SharedLibrary.LocalHttpBrowserSecurity = GetLocalHttpSecurity()

        If Not security.IsBrowserExtensionOrigin(origin) Then Return System.String.Empty

        Dim normalizedOrigin As System.String = System.String.Empty
        If Not SharedLibrary.SharedLibrary.LocalHttpBrowserSecurity.TryNormalizeOrigin(origin, normalizedOrigin) Then
            Return System.String.Empty
        End If

        If security.IsApprovedExternalOrigin(normalizedOrigin) Then Return normalizedOrigin

        Await _localHttpApprovalGate.WaitAsync().ConfigureAwait(False)
        Try
            If security.IsApprovedExternalOrigin(normalizedOrigin) Then Return normalizedOrigin

            Dim approved As System.Boolean = Await SwitchToUi(
                Function() As System.Boolean
                    Dim message As System.String =
                        "A browser extension wants to send content to Red Ink for Word." & vbCrLf & vbCrLf &
                        "Extension origin: " & normalizedOrigin & vbCrLf & vbCrLf &
                        "Approve only if this is the Red Ink browser extension you installed."

                    Dim result As System.Int32 =
                        SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                            message,
                            "Authorize",
                            "Deny",
                            "Red Ink - Authorize Browser Extension")

                    Return result = 1
                End Function).ConfigureAwait(False)

            If Not approved Then Return System.String.Empty
            If Not security.ApproveExternalOrigin(normalizedOrigin) Then Return System.String.Empty

            Return normalizedOrigin
        Finally
            _localHttpApprovalGate.Release()
        End Try
    End Function

    Private Sub AddExplicitCorsOrigin(
        ByVal res As System.Net.HttpListenerResponse,
        ByVal origin As System.String
    )
        If res Is Nothing OrElse System.String.IsNullOrWhiteSpace(origin) Then Return
        res.AddHeader("Access-Control-Allow-Origin", origin)
        res.AddHeader("Vary", "Origin")
    End Sub

End Class
