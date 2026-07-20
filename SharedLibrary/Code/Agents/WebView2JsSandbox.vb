' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: WebView2JsSandbox.vb
' Purpose: Runs untrusted JavaScript inside a hidden, pooled WebView2 instance.
'          Pool size = 1 (reused across calls; per-call lifecycle is semaphore +
'          script execution; underlying control is long-lived).
'
' Threading & Security:
'  - WebView2 lives on host UI (STA) thread; Initialize(...) called once by host.
'  - Network DISABLED by default (all WebResourceRequested => 403).
'    Set allow_network=true on tool to permit fetch().
'  - DevTools/context-menus/host-objects disabled.
'  - User code wrapped: console.log captured, exceptions become error, timeouts
'    return timeout error.
'  - Return value JSON-stringified by wrapper, then unwrapped here.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports System.Net
Imports System.Net.Sockets

Namespace Agents

    Public NotInheritable Class WebView2JsSandbox

        Private Sub New()
        End Sub

        Private Shared ReadOnly _gate As New SemaphoreSlim(1, 1)
        Private Shared _uiSync As SynchronizationContext = Nothing
        Private Shared _userDataDir As String = Nothing
        Private Shared _hostForm As Form = Nothing
        Private Shared _wv As WebView2 = Nothing
        Private Shared _initialized As Boolean = False

        ''' <summary>Host call (UI thread). syncContext must be a WindowsFormsSynchronizationContext.</summary>
        Public Shared Sub Initialize(syncContext As SynchronizationContext, Optional userDataDir As String = Nothing)
            _uiSync = syncContext
            _userDataDir = If(String.IsNullOrWhiteSpace(userDataDir),
                              SharedLibrary.SharedMethods.GetWebView2UserDataFolder(),
                              userDataDir)
        End Sub

        Public Shared ReadOnly Property IsConfigured As Boolean
            Get
                Return _uiSync IsNot Nothing
            End Get
        End Property

        ''' <summary>
        ''' Runs the supplied JS source. Returns the inner JSON envelope produced by the wrapper:
        '''   { "ok": true|false, "result": ..., "logs": [...], "error"?: "..." }
        ''' On timeout returns { "error": "timeout", "timeout_ms": N }.
        ''' </summary>
        Public Shared Async Function RunAsync(code As String,
                                              Optional timeoutMs As Integer = 15000,
                                              Optional allowNetwork As Boolean = False,
                                              Optional navigateUrl As String = "",
                                              Optional waitAfterLoadMs As Integer = 1500,
                                              Optional waitForSelector As String = "",
                                              Optional cancellationToken As CancellationToken = Nothing) As Task(Of String)

            Dim resultJson As String = ""

            If String.IsNullOrWhiteSpace(code) Then
                Return JsonConvert.SerializeObject(New With {Key .error = "missing_code"})
            End If

            If _uiSync Is Nothing Then
                Return JsonConvert.SerializeObject(New With {
                    Key .error = "sandbox_uninitialized",
                    Key .message = "WebView2JsSandbox.Initialize must be called by the host at startup."})
            End If

            If timeoutMs < 500 Then timeoutMs = 500
            If timeoutMs > 120000 Then timeoutMs = 120000

            If waitAfterLoadMs < 0 Then waitAfterLoadMs = 0
            If waitAfterLoadMs > 30000 Then waitAfterLoadMs = 30000

            navigateUrl = If(navigateUrl, "").Trim()
            waitForSelector = If(waitForSelector, "").Trim()

            If navigateUrl <> "" Then
                If Not allowNetwork Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .error = "network_required",
                        Key .message = "navigate_url requires allow_network=true."
                    })
                End If

                If Not IsSafeNavigationUrl(navigateUrl) Then
                    Return JsonConvert.SerializeObject(New With {
                        Key .error = "unsafe_navigate_url",
                        Key .message = "The requested navigate_url was blocked by sandbox policy."
                    })
                End If
            End If

            Await _gate.WaitAsync(cancellationToken).ConfigureAwait(False)

            Try
                Await EnsureInitialized(allowNetwork).ConfigureAwait(False)

                Dim prepareError As String = Await RunOnUiAsync(Of String)(
                    Function() PrepareDocumentAsync(
                        navigateUrl,
                        waitAfterLoadMs,
                        waitForSelector,
                        timeoutMs,
                        cancellationToken)).ConfigureAwait(False)

                If Not String.IsNullOrWhiteSpace(prepareError) Then
                    resultJson = JsonConvert.SerializeObject(New With {
                        Key .error = "navigation_failed",
                        Key .message = prepareError
                    })
                Else
                    Dim wrapped = BuildWrappedScript(code)

                    Dim execTask = RunOnUiAsync(Of String)(Function() ExecuteScriptAwaitPromiseAsync(wrapped))
                    Dim delay = Task.Delay(timeoutMs, cancellationToken)
                    Dim winner = Await Task.WhenAny(execTask, delay).ConfigureAwait(False)

                    If winner Is delay Then
                        resultJson = JsonConvert.SerializeObject(New With {
                            Key .error = "timeout",
                            Key .timeout_ms = timeoutMs
                        })
                    Else
                        Dim inner As String = Await execTask.ConfigureAwait(False)

                        If String.IsNullOrWhiteSpace(inner) Then
                            inner = JsonConvert.SerializeObject(New With {
                                Key .ok = True,
                                Key .result = CType(Nothing, Object)
                            })
                        End If

                        resultJson = inner
                    End If
                End If

            Catch oce As OperationCanceledException
                resultJson = JsonConvert.SerializeObject(New With {Key .error = "cancelled"})
            Catch ex As Exception
                resultJson = JsonConvert.SerializeObject(New With {Key .error = "js_run_failed", Key .message = ex.Message})
            End Try

            Try
                Await RunOnUiAsync(Function() ResetDocumentAsync()).ConfigureAwait(False)
            Catch
            End Try

            _gate.Release()
            Return resultJson
        End Function

        Private Shared Async Function PrepareDocumentAsync(navigateUrl As String,
                                                           waitAfterLoadMs As Integer,
                                                           waitForSelector As String,
                                                           timeoutMs As Integer,
                                                           cancellationToken As CancellationToken) As Task(Of String)
            If _wv Is Nothing OrElse _wv.CoreWebView2 Is Nothing Then
                Return "Sandbox browser is not initialized."
            End If

            If String.IsNullOrWhiteSpace(navigateUrl) Then
                Await ResetDocumentAsync().ConfigureAwait(True)
                Return ""
            End If

            Return Await NavigateAndWaitAsync(
                navigateUrl,
                waitAfterLoadMs,
                waitForSelector,
                Math.Min(timeoutMs, 45000),
                cancellationToken).ConfigureAwait(True)
        End Function

        Private Shared Async Function ResetDocumentAsync() As Task
            If _wv Is Nothing OrElse _wv.CoreWebView2 Is Nothing Then
                Return
            End If

            Dim tcs As New TaskCompletionSource(Of Boolean)()
            Dim handler As EventHandler(Of CoreWebView2NavigationCompletedEventArgs) = Nothing

            handler =
                Sub(sender, e)
                    Try
                        RemoveHandler _wv.CoreWebView2.NavigationCompleted, handler
                    Catch
                    End Try

                    tcs.TrySetResult(True)
                End Sub

            AddHandler _wv.CoreWebView2.NavigationCompleted, handler
            _wv.CoreWebView2.NavigateToString("<!doctype html><html><head><meta charset=""utf-8""></head><body></body></html>")

            Dim completed = Await Task.WhenAny(tcs.Task, Task.Delay(2000)).ConfigureAwait(True)

            If completed IsNot tcs.Task Then
                Try
                    RemoveHandler _wv.CoreWebView2.NavigationCompleted, handler
                Catch
                End Try
            End If
        End Function

        Private Shared Async Function NavigateAndWaitAsync(url As String,
                                                           waitAfterLoadMs As Integer,
                                                           waitForSelector As String,
                                                           timeoutMs As Integer,
                                                           cancellationToken As CancellationToken) As Task(Of String)
            Dim tcs As New TaskCompletionSource(Of String)()
            Dim handler As EventHandler(Of CoreWebView2NavigationCompletedEventArgs) = Nothing

            handler =
                Sub(sender, e)
                    Try
                        RemoveHandler _wv.CoreWebView2.NavigationCompleted, handler
                    Catch
                    End Try

                    If e.IsSuccess Then
                        tcs.TrySetResult("")
                    Else
                        tcs.TrySetResult("Navigation failed: " & e.WebErrorStatus.ToString())
                    End If
                End Sub

            AddHandler _wv.CoreWebView2.NavigationCompleted, handler
            _wv.CoreWebView2.Navigate(url)

            Dim winner = Await Task.WhenAny(
                tcs.Task,
                Task.Delay(Math.Max(1000, timeoutMs), cancellationToken)).ConfigureAwait(True)

            If winner IsNot tcs.Task Then
                Try
                    RemoveHandler _wv.CoreWebView2.NavigationCompleted, handler
                Catch
                End Try

                Return "Navigation timed out."
            End If

            Dim navError As String = Await tcs.Task.ConfigureAwait(True)
            If Not String.IsNullOrWhiteSpace(navError) Then
                Return navError
            End If

            If waitAfterLoadMs > 0 Then
                Await Task.Delay(waitAfterLoadMs, cancellationToken).ConfigureAwait(True)
            End If

            If Not String.IsNullOrWhiteSpace(waitForSelector) Then
                Dim found As Boolean = Await WaitForSelectorAsync(
                    waitForSelector,
                    Math.Min(timeoutMs, 15000),
                    cancellationToken).ConfigureAwait(True)

                If Not found Then
                    Return "Selector did not appear before timeout: " & waitForSelector
                End If
            End If

            Return ""
        End Function

        Private Shared Async Function WaitForSelectorAsync(selector As String,
                                                           timeoutMs As Integer,
                                                           cancellationToken As CancellationToken) As Task(Of Boolean)
            If String.IsNullOrWhiteSpace(selector) Then
                Return True
            End If

            Dim selectorLiteral As String = JsonConvert.SerializeObject(selector)

            Dim script As String = <![CDATA[
(function() {
    var selector = __SELECTOR__;

    function rootHasSelector(root) {
        if (!root) return false;

        try {
            if (root.querySelector && root.querySelector(selector)) {
                return true;
            }
        } catch (e) {
            return false;
        }

        try {
            var all = root.querySelectorAll ? root.querySelectorAll('*') : [];
            for (var i = 0; i < all.length; i++) {
                var el = all[i];
                if (el && el.shadowRoot && rootHasSelector(el.shadowRoot)) {
                    return true;
                }
            }
        } catch (e) {
        }

        return false;
    }

    return rootHasSelector(document);
})();
]]>.Value.Replace("__SELECTOR__", selectorLiteral)

            Dim started As DateTime = DateTime.UtcNow

            Do
                cancellationToken.ThrowIfCancellationRequested()

                Try
                    Dim raw As String = Await _wv.CoreWebView2.ExecuteScriptAsync(script).ConfigureAwait(True)
                    If String.Equals(If(raw, "").Trim(), "true", StringComparison.OrdinalIgnoreCase) Then
                        Return True
                    End If
                Catch
                End Try

                If (DateTime.UtcNow - started).TotalMilliseconds >= timeoutMs Then
                    Exit Do
                End If

                Await Task.Delay(250, cancellationToken).ConfigureAwait(True)
            Loop

            Return False
        End Function

        Private Shared Function IsSafeNavigationUrl(url As String) As Boolean
            Try
                Dim uriResult As Uri = Nothing
                If Not Uri.TryCreate(url, UriKind.Absolute, uriResult) Then
                    Return False
                End If

                If uriResult.Scheme <> Uri.UriSchemeHttp AndAlso uriResult.Scheme <> Uri.UriSchemeHttps Then
                    Return False
                End If

                If uriResult.IsLoopback Then
                    Return False
                End If

                If Not String.IsNullOrWhiteSpace(uriResult.UserInfo) Then
                    Return False
                End If

                Dim host As String = If(uriResult.Host, "").Trim().ToLowerInvariant()
                If host = "" Then
                    Return False
                End If

                If host = "localhost" OrElse
                   host.EndsWith(".local", StringComparison.OrdinalIgnoreCase) OrElse
                   host.EndsWith(".internal", StringComparison.OrdinalIgnoreCase) OrElse
                   host.EndsWith(".home", StringComparison.OrdinalIgnoreCase) Then
                    Return False
                End If

                Dim literalIp As IPAddress = Nothing
                If IPAddress.TryParse(host, literalIp) Then
                    Return Not IsPrivateIpAddress(literalIp)
                End If

                Try
                    For Each resolvedAddress In Dns.GetHostAddresses(uriResult.DnsSafeHost)
                        If IsPrivateIpAddress(resolvedAddress) Then
                            Return False
                        End If
                    Next
                Catch
                    ' If DNS resolution fails here, allow the normal navigation attempt to decide reachability.
                End Try

                Return True
            Catch
                Return False
            End Try
        End Function

        Private Shared Function IsPrivateIpAddress(address As IPAddress) As Boolean
            If address Is Nothing Then
                Return True
            End If

            If IPAddress.IsLoopback(address) Then
                Return True
            End If

            Dim bytes = address.GetAddressBytes()

            If address.AddressFamily = AddressFamily.InterNetwork Then
                If bytes.Length <> 4 Then
                    Return True
                End If

                If bytes(0) = 10 Then Return True
                If bytes(0) = 127 Then Return True
                If bytes(0) = 169 AndAlso bytes(1) = 254 Then Return True
                If bytes(0) = 172 AndAlso bytes(1) >= 16 AndAlso bytes(1) <= 31 Then Return True
                If bytes(0) = 192 AndAlso bytes(1) = 168 Then Return True
                If bytes(0) = 100 AndAlso bytes(1) >= 64 AndAlso bytes(1) <= 127 Then Return True
                If bytes(0) = 0 Then Return True

                Return False
            End If

            If address.AddressFamily = AddressFamily.InterNetworkV6 Then
                If address.IsIPv6LinkLocal OrElse address.IsIPv6SiteLocal Then
                    Return True
                End If

                If bytes.Length = 16 AndAlso (bytes(0) And &HFE) = &HFC Then
                    Return True
                End If

                Return False
            End If

            Return True
        End Function


        ' --------------------------------------------------------------- wrapper

        Private Shared Function BuildWrappedScript(userCode As String) As String
            Dim literal = JsonConvert.SerializeObject(userCode)
            Return _
"(async () => {" & vbLf &
"  const __logs = [];" & vbLf &
"  console.log = (...a) => { try { __logs.push(a.map(x => x===undefined?'undefined':(x===null?'null':(typeof x==='object'?JSON.stringify(x):String(x)))).join(' ')); } catch(_) {} };" & vbLf &
"  console.error = console.log; console.warn = console.log;" & vbLf &
"  try {" & vbLf &
"    const __fn = new Function('return (async () => { ' + " & literal & " + ' \n })();');" & vbLf &
"    const __r = await __fn();" & vbLf &
"    return JSON.stringify({ ok: true, result: __r === undefined ? null : __r, logs: __logs });" & vbLf &
"  } catch (e) {" & vbLf &
"    return JSON.stringify({ ok: false, error: (e && e.message) ? String(e.message) : String(e), logs: __logs });" & vbLf &
"  }" & vbLf &
"})()"
        End Function

        ' --------------------------------------------------------------- lifecycle

        Private Shared Async Function ExecuteScriptAwaitPromiseAsync(script As String) As Task(Of String)
            Try
                Dim payload As String = New JObject(
            New JProperty("expression", script),
            New JProperty("awaitPromise", True),
            New JProperty("returnByValue", True)
        ).ToString(Formatting.None)

                Dim raw As String = Await _wv.CoreWebView2.CallDevToolsProtocolMethodAsync(
            "Runtime.evaluate",
            payload).ConfigureAwait(True)

                Dim root As JObject = JObject.Parse(raw)

                Dim exceptionToken As JToken = root("exceptionDetails")
                If exceptionToken IsNot Nothing Then
                    Dim message As String = If(
                exceptionToken.SelectToken("exception.description")?.ToString(),
                exceptionToken.SelectToken("text")?.ToString())

                    If String.IsNullOrWhiteSpace(message) Then
                        message = exceptionToken.ToString(Formatting.None)
                    End If

                    Return JsonConvert.SerializeObject(New With {
                Key .ok = False,
                Key .error = message
            })
                End If

                Dim valueToken As JToken = root.SelectToken("result.value")
                If valueToken IsNot Nothing Then
                    If valueToken.Type = JTokenType.String Then
                        Return valueToken.ToString()
                    End If

                    Return valueToken.ToString(Formatting.None)
                End If

                Dim unserializableToken As JToken = root.SelectToken("result.unserializableValue")
                If unserializableToken IsNot Nothing Then
                    Return JsonConvert.SerializeObject(New With {
                Key .ok = True,
                Key .result = unserializableToken.ToString()
            })
                End If

                Return JsonConvert.SerializeObject(New With {
            Key .ok = True,
            Key .result = CType(Nothing, Object)
        })

            Catch ex As Exception
                Return JsonConvert.SerializeObject(New With {
            Key .ok = False,
            Key .error = "js_run_eval_failed",
            Key .message = ex.Message
        })
            End Try
        End Function

        Private Shared Async Function EnsureInitialized(allowNetwork As Boolean) As Task
            If _initialized Then
                Await RunOnUiAsync(Sub() ConfigureNetworkPolicy(allowNetwork)).ConfigureAwait(False)
                Return
            End If
            Await RunOnUiAsync(Async Function() As Task
                                   If _hostForm Is Nothing Then
                                       _hostForm = New Form() With {
                                           .Text = "RedInk JS Sandbox",
                                           .FormBorderStyle = FormBorderStyle.None,
                                           .ShowInTaskbar = False,
                                           .StartPosition = FormStartPosition.Manual,
                                           .Location = New Point(-32000, -32000),
                                           .Size = New Size(2, 2),
                                           .Opacity = 0.0
                                       }
                                       AddHandler _hostForm.FormClosing, Sub(s, e) e.Cancel = True
                                       _hostForm.Show()
                                       _hostForm.Hide()
                                   End If
                                   If _wv Is Nothing Then
                                       _wv = New WebView2() With {.Dock = DockStyle.Fill}
                                       _hostForm.Controls.Add(_wv)
                                       Dim env As CoreWebView2Environment = Nothing
                                       Dim envOk As Boolean = True
                                       Try
                                           env = Await CoreWebView2Environment.CreateAsync(Nothing, _userDataDir).ConfigureAwait(True)
                                       Catch
                                           envOk = False
                                       End Try
                                       If Not envOk Then
                                           env = Await CoreWebView2Environment.CreateAsync().ConfigureAwait(True)
                                       End If
                                       Await _wv.EnsureCoreWebView2Async(env).ConfigureAwait(True)
                                       Await _wv.EnsureCoreWebView2Async(env).ConfigureAwait(True)
                                       Dim s = _wv.CoreWebView2.Settings
                                       s.AreDevToolsEnabled = False
                                       s.AreDefaultContextMenusEnabled = False
                                       s.IsStatusBarEnabled = False
                                       s.IsZoomControlEnabled = False
                                       s.IsBuiltInErrorPageEnabled = False
                                       ConfigureNetworkPolicy(allowNetwork)
                                       AddHandler _wv.CoreWebView2.ProcessFailed, AddressOf OnSandboxProcessFailed
                                       _wv.CoreWebView2.NavigateToString("<!doctype html><html><head><meta charset=""utf-8""><title>sandbox</title></head><body></body></html>")
                                   End If
                                   _initialized = True
                               End Function).ConfigureAwait(False)
        End Function

        ''' <summary>
        ''' Handles a WebView2 render/browser process crash. Tears down the pooled control so the
        ''' next RunAsync call re-initializes cleanly instead of blocking forever on a dead instance.
        ''' </summary>
        Private Shared Sub OnSandboxProcessFailed(sender As Object, e As CoreWebView2ProcessFailedEventArgs)
            SharedLibrary.SharedMethods.LogWebView2ProcessFailed("JsSandbox", e.ProcessFailedKind.ToString(), e.ExitCode.ToString())
            Try
                If _wv IsNot Nothing Then
                    If _hostForm IsNot Nothing AndAlso _hostForm.Controls.Contains(_wv) Then
                        _hostForm.Controls.Remove(_wv)
                    End If
                    _wv.Dispose()
                End If
            Catch
            End Try
            _wv = Nothing
            _initialized = False
        End Sub

        Private Shared Sub ConfigureNetworkPolicy(allow As Boolean)
            Try
                RemoveHandler _wv.CoreWebView2.WebResourceRequested, AddressOf BlockAll
            Catch
            End Try
            Try
                _wv.CoreWebView2.RemoveWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All)
            Catch
            End Try
            If Not allow Then
                Try
                    _wv.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All)
                    AddHandler _wv.CoreWebView2.WebResourceRequested, AddressOf BlockAll
                Catch
                End Try
            End If
        End Sub

        Private Shared Sub BlockAll(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
            Try
                Dim env = _wv.CoreWebView2.Environment
                e.Response = env.CreateWebResourceResponse(Nothing, 403, "Forbidden", "Content-Type: text/plain")
            Catch
            End Try
        End Sub

        ' --------------------------------------------------------------- UI marshalling

        Private Shared Function RunOnUiAsync(action As Action) As Task
            Dim tcs As New TaskCompletionSource(Of Boolean)()
            _uiSync.Post(Sub()
                             Try : action() : tcs.SetResult(True)
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

        Private Shared Function RunOnUiAsync(Of T)(func As Func(Of Task(Of T))) As Task(Of T)
            Dim tcs As New TaskCompletionSource(Of T)()
            _uiSync.Post(Async Sub()
                             Try : tcs.SetResult(Await func().ConfigureAwait(True))
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

        Private Shared Function RunOnUiAsync(func As Func(Of Task)) As Task
            Dim tcs As New TaskCompletionSource(Of Boolean)()
            _uiSync.Post(Async Sub()
                             Try : Await func().ConfigureAwait(True) : tcs.SetResult(True)
                             Catch ex As Exception : tcs.SetException(ex)
                             End Try
                         End Sub, Nothing)
            Return tcs.Task
        End Function

    End Class

End Namespace
