' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebExtension.vb
' Purpose: Provides HTTP listener functionality for Word to receive commands from
'          external applications (primarily browser extensions) via localhost.
'
' Architecture:
'  - HTTP Listener: Starts an async HttpListener on http://127.0.0.1:12334/ to
'    receive JSON-formatted commands from external sources.
'  - Threading Model: Uses async/await pattern throughout; UI operations marshaled
'    to main thread via SwitchToUi helpers using mainThreadControl.Invoke.
'  - Command Dispatch: ProcessRequestInAddIn parses incoming JSON body and routes
'    to appropriate handlers based on "Command" field.
'  - Supported Commands:
'    * "redink_sendtoword": Inserts text and URL into current Word selection/cursor.
'  - Error Recovery: Tracks consecutive failures; restarts listener after 10
'    consecutive errors with 5-second delay.
'  - CORS Support: Handles OPTIONS preflight requests for cross-origin scenarios.
'  - Lifecycle: StartupHttpListener/ShutdownHttpListener manage listener state;
'    isShuttingDown flag prevents restart during add-in shutdown.
'  - COM Hygiene: Explicitly releases COM objects (Selection) to prevent leaks.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Diagnostics
Imports System.Windows.Forms

Partial Public Class ThisAddIn

    Private httpListener As System.Net.HttpListener
    Private listenerTask As System.Threading.Tasks.Task
    Private isShuttingDown As Boolean = False

    Private Function IsEdgeExtensionReceiverEnabled() As System.Boolean
        Return INI_WebServerBlock <> 1 AndAlso INI_WebServerBlock <> 3
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Executes an action on the UI thread and waits for completion.
    ''' </summary>
    ''' <param name="uiAction">Action to execute on the main thread.</param>
    ''' <returns>Task that completes when the action finishes or faults.</returns>
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(uiAction As System.Action) _
        As System.Threading.Tasks.Task

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of Object)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                uiAction.Invoke()
                tcs.SetResult(Nothing)
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Executes a function on the UI thread and waits for its result.
    ''' </summary>
    ''' <typeparam name="T">Return type of the function.</typeparam>
    ''' <param name="uiFunc">Function to execute on the main thread.</param>
    ''' <returns>Task containing the function's result or exception.</returns>
    '───────────────────────────────────────────────────────────────────────────
    Private Function SwitchToUi(Of T)(uiFunc As System.Func(Of T)) _
        As System.Threading.Tasks.Task(Of T)

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of T)()

        mainThreadControl.Invoke(New MethodInvoker(
        Sub()
            Try
                tcs.SetResult(uiFunc.Invoke())
            Catch ex As System.Exception
                tcs.SetException(ex)
            End Try
        End Sub))

        Return tcs.Task
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Starts the HTTP listener in a fire-and-forget manner.
    ''' </summary>
    '───────────────────────────────────────────────────────────────────────────
    Private Sub StartupHttpListener(WebServerBlock As Integer)
        If WebServerBlock = 1 OrElse WebServerBlock = 3 Then Return
        listenerTask = StartHttpListener()
    End Sub

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Stops the HTTP listener and releases resources.
    ''' </summary>
    '───────────────────────────────────────────────────────────────────────────
    Private Sub ShutdownHttpListener()
        isShuttingDown = True
        If httpListener IsNot Nothing AndAlso httpListener.IsListening Then
            httpListener.Stop()
            httpListener.Close()
        End If
    End Sub

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Main HTTP listener loop. Accepts incoming connections on port 12334,
    ''' handles requests, and recovers from failures automatically.
    ''' </summary>
    ''' <returns>Task that completes when listener shuts down.</returns>
    '───────────────────────────────────────────────────────────────────────────
    Private Async Function StartHttpListener() As System.Threading.Tasks.Task
        Const prefix As String = "http://127.0.0.1:12334/"
        Dim consecutiveFailures As Integer = 0

        While Not isShuttingDown
            Try
                ' Ensure listener exists and is running
                If httpListener Is Nothing Then
                    httpListener = New System.Net.HttpListener()
                    httpListener.Prefixes.Add(prefix)
                    httpListener.Start()
                    Debug.WriteLine("HttpListener started.")
                ElseIf Not httpListener.IsListening Then
                    httpListener.Close()
                    httpListener = Nothing
                    Continue While
                End If

                ' Wait for one incoming request
                Dim ctx As System.Net.HttpListenerContext =
                Await httpListener.GetContextAsync().ConfigureAwait(False)

                ' Handle the request (fire-and-forget)
                Call HandleHttpRequest(ctx) _
                .ContinueWith(
                    Sub(t)
                        If t.IsFaulted AndAlso t.Exception IsNot Nothing Then
                            Debug.WriteLine("HandleHttpRequest error: " &
                                            t.Exception.GetBaseException().Message)
                        End If
                    End Sub,
                    System.Threading.Tasks.TaskScheduler.Default)

                consecutiveFailures = 0
            Catch ex As System.ObjectDisposedException
                consecutiveFailures += 1
            Catch ex As System.Exception
                consecutiveFailures += 1
                Debug.WriteLine("Listener error: " & ex.Message)
            End Try

            ' Recycle after too many consecutive errors
            If consecutiveFailures >= 10 AndAlso Not isShuttingDown Then
                Debug.WriteLine("Restarting HttpListener after 10 failures.")
                Try
                    If httpListener IsNot Nothing Then httpListener.Close()
                Catch
                End Try
                httpListener = Nothing
                consecutiveFailures = 0
                Await System.Threading.Tasks.Task.Delay(5000).ConfigureAwait(False)
            End If
        End While
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Handles a single HTTP request: processes CORS preflight, reads body,
    ''' dispatches to command handler, and sends response.
    ''' </summary>
    ''' <param name="ctx">The HTTP listener context containing request and response.</param>
    ''' <returns>Task that completes when response is sent.</returns>
    '───────────────────────────────────────────────────────────────────────────
    Private Async Function HandleHttpRequest(
        ctx As System.Net.HttpListenerContext) _
        As System.Threading.Tasks.Task

        Dim req As System.Net.HttpListenerRequest = ctx.Request
        Dim res As System.Net.HttpListenerResponse = ctx.Response

        If Not IsEdgeExtensionReceiverEnabled() Then
            Dim msgBytes() As System.Byte =
                System.Text.Encoding.UTF8.GetBytes("The Edge extension receiver is disabled by WebServerBlock.")

            res.StatusCode = 403
            res.ContentLength64 = msgBytes.LongLength
            res.ContentType = "text/plain; charset=utf-8"
            res.AddHeader("Access-Control-Allow-Origin", "*")

            Using os As System.IO.Stream = res.OutputStream
                Await os.WriteAsync(msgBytes, 0, msgBytes.Length).ConfigureAwait(False)
            End Using

            res.Close()
            Return
        End If

        ' CORS pre-flight
        If req.HttpMethod = "OPTIONS" Then
            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
            res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
            res.StatusCode = 204 : res.Close() : Return
        End If

        ' Read body (if any)
        Dim body As String = ""
        If req.HasEntityBody Then
            Using rdr As New IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8)
                body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
            End Using
        End If

        ' Dispatch to add-in logic
        Dim responseText As String =
        Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)

        ' Send response
        Dim buf = System.Text.Encoding.UTF8.GetBytes(responseText)
        res.ContentLength64 = buf.Length
        res.ContentType = "text/plain; charset=utf-8"
        res.AddHeader("Access-Control-Allow-Origin", "*")
        Using os = res.OutputStream
            Await os.WriteAsync(buf, 0, buf.Length).ConfigureAwait(False)
        End Using
        res.Close()
    End Function

    '───────────────────────────────────────────────────────────────────────────
    ''' <summary>
    ''' Parses incoming JSON command and routes to appropriate handler.
    ''' Currently supports "redink_sendtoword" command which inserts text and URL
    ''' into the active Word document at the current cursor/selection.
    ''' </summary>
    ''' <param name="body">JSON body containing Command, Text, and URL fields.</param>
    ''' <param name="rawUrl">Raw request URL (currently unused).</param>
    ''' <returns>Response string to send back to client (empty for most commands).</returns>
    '───────────────────────────────────────────────────────────────────────────
    Private Async Function ProcessRequestInAddIn(
        body As String,
        rawUrl As String) _
        As System.Threading.Tasks.Task(Of String)

        ' Guard clause: empty body
        If String.IsNullOrWhiteSpace(body) Then Return ""

        Dim j = Newtonsoft.Json.Linq.JObject.Parse(body)
        Dim cmd = j("Command")?.ToString()
        Dim textBody = j("Text")?.ToString()
        Dim sourceUrl = j("URL")?.ToString()

        Select Case cmd
            Case "redink_sendtoword"
                If String.IsNullOrWhiteSpace(textBody) Then Return ""

                ' Everything that touches Word must run on the UI thread
                Await SwitchToUi(Sub()

                                     Dim wdApp As Microsoft.Office.Interop.Word.Application =
                    Globals.ThisAddIn.Application

                                     Dim sel As Microsoft.Office.Interop.Word.Selection = wdApp.Selection

                                     wdApp.ScreenUpdating = False

                                     ' Sanitize inputs and insert without using TypeText
                                     Dim safeBody As String = If(textBody, String.Empty).Replace(ChrW(0), String.Empty)
                                     Dim safeUrl As String = If(sourceUrl, String.Empty).Replace(ChrW(0), String.Empty)
                                     Dim finalText As String = safeBody & " (" & safeUrl & ")"

                                     Dim rng As Word.Range = sel.Range.Duplicate
                                     If rng.Start = rng.End Then
                                         rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)
                                         rng.InsertAfter(finalText)
                                     Else
                                         rng.Text = finalText
                                     End If

                                     sel.SetRange(rng.End, rng.End)
                                     wdApp.ScreenUpdating = True

                                     ' Release COM objects explicitly
                                     System.Runtime.InteropServices.Marshal.ReleaseComObject(sel)
                                 End Sub)

                Return ""
        End Select

        Return ""
    End Function

End Class
