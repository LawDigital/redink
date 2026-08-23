' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.WebExtension.Security.vb
' Purpose:
'   Host-specific wiring for the shared localhost HTTP security layer, including
'   persistent browser credentials, CSRF/origin validation, extension-origin approval,
'   custom Red Ink authorization dialogs, branded local pairing/denial feedback,
'   and shared bounded-request enforcement.
'
' Regression invariants:
'   - The local chat remains reachable through the existing / and /inky URLs.
'   - Existing /inky/api command names and payloads remain unchanged.
'   - Existing Chromium browser-extension POST payloads to /redink remain unchanged.
'   - Persistence is optional: missing My.Settings entries fall back to process-memory
'     state, never to unauthenticated access.
' =============================================================================

Option Explicit On
Option Strict On

Partial Public Class ThisAddIn

    Private Const InkyPairRoute As System.String = "/inky/auth/pair"
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
        Try
            AppendInkyServerLog("LocalHttpSecurity: " & If(message, System.String.Empty))
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("[LocalHttpSecurity] " & If(message, System.String.Empty))
        End Try
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

    Private Function EnsureLocalLoopbackRequest(
        ByVal req As System.Net.HttpListenerRequest,
        ByVal res As System.Net.HttpListenerResponse,
        ByVal requestId As System.String
    ) As System.Boolean
        If GetLocalHttpSecurity().IsLoopbackRequest(req) Then Return True

        Dim payload() As System.Byte =
            System.Text.Encoding.UTF8.GetBytes("Forbidden: localhost access required.")

        SendBufferedHttpResponse(
            res,
            403,
            "text/plain; charset=utf-8",
            payload,
            requestId,
            "non-loopback-rejected",
            addCors:=False)

        Return False
    End Function

    Private Function TryAuthorizeLocalChatRequest(
        ByVal req As System.Net.HttpListenerRequest,
        ByVal res As System.Net.HttpListenerResponse,
        ByVal requestId As System.String,
        ByVal requireCsrf As System.Boolean,
        ByRef browserCredential As System.String
    ) As System.Boolean
        browserCredential = System.String.Empty

        If Not GetLocalHttpSecurity().TryValidateBrowserCredential(req.Headers("Cookie"), browserCredential) Then
            Dim payload() As System.Byte =
                System.Text.Encoding.UTF8.GetBytes("Unauthorized browser. Open the local chat page and authorize this browser.")

            SendBufferedHttpResponse(
                res,
                401,
                "text/plain; charset=utf-8",
                payload,
                requestId,
                "local-chat-auth-required",
                addCors:=False)

            Return False
        End If

        Dim origin As System.String = If(req.Headers("Origin"), System.String.Empty)
        If Not GetLocalHttpSecurity().IsSameOrigin(origin, req.Url) Then
            Dim payload() As System.Byte =
                System.Text.Encoding.UTF8.GetBytes("Forbidden origin.")

            SendBufferedHttpResponse(
                res,
                403,
                "text/plain; charset=utf-8",
                payload,
                requestId,
                "local-chat-origin-rejected",
                addCors:=False)

            Return False
        End If

        If requireCsrf Then
            Dim csrfToken As System.String = If(req.Headers("X-RedInk-CSRF"), System.String.Empty)
            If Not GetLocalHttpSecurity().ValidateCsrfToken(browserCredential, csrfToken) Then
                Dim payload() As System.Byte =
                    System.Text.Encoding.UTF8.GetBytes("Forbidden request token.")

                SendBufferedHttpResponse(
                    res,
                    403,
                    "text/plain; charset=utf-8",
                    payload,
                    requestId,
                    "local-chat-csrf-rejected",
                    addCors:=False)

                Return False
            End If
        End If

        Return True
    End Function

    Private Async Function HandleLocalBrowserPairRequestAsync(
        ByVal req As System.Net.HttpListenerRequest,
        ByVal res As System.Net.HttpListenerResponse,
        ByVal requestId As System.String
    ) As System.Threading.Tasks.Task
        If Not GetLocalHttpSecurity().IsSameOrigin(If(req.Headers("Origin"), System.String.Empty), req.Url) Then
            Dim forbiddenBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Forbidden origin.")
            SendBufferedHttpResponse(
                res,
                403,
                "text/plain; charset=utf-8",
                forbiddenBytes,
                requestId,
                "pair-origin-rejected",
                addCors:=False)
            Return
        End If

        Await _localHttpApprovalGate.WaitAsync().ConfigureAwait(False)
        Try
            Dim existingCredential As System.String = System.String.Empty
            If GetLocalHttpSecurity().TryValidateBrowserCredential(req.Headers("Cookie"), existingCredential) Then
                res.AppendHeader("Set-Cookie", GetLocalHttpSecurity().CreateAuthCookieHeader(existingCredential))
                Dim alreadyBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("{""ok"":true}")
                SendBufferedHttpResponse(
                    res,
                    200,
                    "application/json; charset=utf-8",
                    alreadyBytes,
                    requestId,
                    "pair-already-authorized",
                    addCors:=False)
                Return
            End If

            Dim userAgent As System.String = ClipForInkyServerLog(If(req.UserAgent, System.String.Empty), 240)
            Dim origin As System.String = If(req.Headers("Origin"), System.String.Empty)

            Dim approved As System.Boolean = Await SwitchToUi(
                Function() As System.Boolean
                    Dim message As System.String =
                        "A browser wants to connect to Red Ink Local Chat." & vbCrLf & vbCrLf &
                        "Origin: " & origin & vbCrLf &
                        If(System.String.IsNullOrWhiteSpace(userAgent), System.String.Empty, "Browser: " & userAgent & vbCrLf) & vbCrLf &
                        "Approve only if you opened the Red Ink localhost page yourself."

                    Dim result As System.Int32 =
                        SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox(
                            message,
                            "Authorize",
                            "Deny",
                            "Red Ink - Authorize Browser")

                    Return result = 1
                End Function).ConfigureAwait(False)

            If Not approved Then
                Dim deniedBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Authorization was not granted. Please approve the request in Outlook if you want to use this browser with Red Ink Local Chat.")
                SendBufferedHttpResponse(
                    res,
                    403,
                    "text/plain; charset=utf-8",
                    deniedBytes,
                    requestId,
                    "pair-denied",
                    addCors:=False)
                Return
            End If

            Dim browserCredential As System.String = GetLocalHttpSecurity().IssueBrowserCredential()
            res.AppendHeader("Set-Cookie", GetLocalHttpSecurity().CreateAuthCookieHeader(browserCredential))

            Dim okBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("{""ok"":true}")
            SendBufferedHttpResponse(
                res,
                200,
                "application/json; charset=utf-8",
                okBytes,
                requestId,
                "pair-approved",
                addCors:=False)
        Finally
            _localHttpApprovalGate.Release()
        End Try
    End Function

    Private Function GetLocalBrowserPairingLogoDataUri() As System.String
        Try
            Dim sourceBitmap As System.Drawing.Bitmap =
                SharedLibrary.SharedLibrary.SharedMethods.GetLogoBitmap(
                    SharedLibrary.SharedLibrary.SharedMethods.LogoType.Standard,
                    RedInkLogo:=True)

            If sourceBitmap Is Nothing Then Return System.String.Empty

            Using logoBitmap As New System.Drawing.Bitmap(sourceBitmap)
                Using stream As New System.IO.MemoryStream()
                    logoBitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png)
                    Return "data:image/png;base64," & System.Convert.ToBase64String(stream.ToArray())
                End Using
            End Using
        Catch ex As System.Exception
            LogLocalHttpSecurity("Cannot render Red Ink logo for browser authorization page (" & ex.GetType().Name & ").")
            Return System.String.Empty
        End Try
    End Function

    Private Function BuildLocalBrowserPairingPage() As System.String
        Dim brandName As System.String = If(Not System.String.IsNullOrWhiteSpace(AN), AN, "Red Ink")
        Dim html As New System.Text.StringBuilder()
        Dim logoDataUri As System.String = GetLocalBrowserPairingLogoDataUri()

        html.AppendLine("<!doctype html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<meta name=""viewport"" content=""width=device-width,initial-scale=1"">")
        html.AppendLine("<meta http-equiv=""Content-Security-Policy"" content=""default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; connect-src 'self'; img-src data:"">")
        html.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " - Browser Authorization</title>")
        html.AppendLine("<style>html,body{height:100%;margin:0;font-family:system-ui,Segoe UI,Arial,sans-serif;background:#0b0f14;color:#e8eef6}.wrap{height:100%;display:flex;align-items:center;justify-content:center;padding:24px;box-sizing:border-box}.card{max-width:560px;background:#11161d;border:1px solid #1b2430;border-radius:14px;padding:28px;box-shadow:0 10px 30px rgba(0,0,0,.35)}.logo{display:block;max-width:190px;max-height:72px;width:auto;height:auto;margin:0 0 22px 0}h1{margin-top:0;font-size:1.5rem}p{line-height:1.5;color:#c7d0da}button{background:#e8eef6;color:#11161d;border:0;border-radius:9px;padding:11px 16px;font:inherit;font-weight:650;cursor:pointer}button:disabled{opacity:.6;cursor:wait}.status{margin-top:14px;color:#9aa8b7;font-size:.9rem}</style></head><body>")
        Dim logoHtml As System.String = System.String.Empty
        If Not System.String.IsNullOrWhiteSpace(logoDataUri) Then
            logoHtml = "<img class=""logo"" alt=""Red Ink"" src=""" & System.Net.WebUtility.HtmlEncode(logoDataUri) & """>"
        End If

        html.AppendLine("<div class=""wrap""><div class=""card"">" & logoHtml & "<h1>Authorize this browser</h1><p>This browser is not yet authorized to control Red Ink Local Chat. Click below, then approve the confirmation shown by Outlook.</p><button id=""authorize"">Authorize browser</button><div id=""status"" class=""status""></div></div></div>")
        html.AppendLine("<script>'use strict';const b=document.getElementById('authorize'),s=document.getElementById('status');b.addEventListener('click',async()=>{b.disabled=true;s.textContent='Waiting for approval in Outlook...';try{const r=await fetch('/inky/auth/pair',{method:'POST',headers:{'Content-Type':'application/json'},body:'{}'});if(r.ok){s.textContent='Authorized. Opening Local Chat...';location.replace('/inky');return;}const t=await r.text();s.textContent=t||'Authorization was not granted.';}catch(e){s.textContent=e.message||'Authorization failed.';}finally{b.disabled=false;}});</script>")
        html.AppendLine("</body></html>")

        Return html.ToString()
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
                        "A browser extension wants to connect to Red Ink for Outlook." & vbCrLf & vbCrLf &
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

    Private Sub AddLocalUiSecurityHeaders(ByVal res As System.Net.HttpListenerResponse)
        If res Is Nothing Then Return
        res.AddHeader("X-Frame-Options", "DENY")
        res.AddHeader("Content-Security-Policy", "frame-ancestors 'none'")
        res.AddHeader("Referrer-Policy", "no-referrer")
    End Sub

End Class
