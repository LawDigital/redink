' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.WebExtension.vb
' Purpose: Provides the local web extension surface (HTML UI + JSON API) for chat
'          and document related LLM operations. Handles HTTP requests, command
'          dispatch, LLM job scheduling, model selection, chat state persistence,
'          file upload support, markdown conversion, and alternate model application.
'
' Architecture:
' - HTTP layer: single HttpListener (prefix /inky) returns HTML UI (GET) and JSON API (POST).
' - Routing constants: InkyUiRoute (HTML), InkyApiRoute (JSON commands).
' - Chat state: persisted per chat id (1/2) in application settings (InkyState + ChatTurn).
' - Job management: each LLM invocation tracked as LlmJob (ConcurrentDictionary jobMap + TTL).
' - Cancellation: per-job CancellationTokenSource plus legacy global llmOperationCts.
' - Alternate models: optional secondary / alternate model application guarded by lock.
' - File handling: optional file upload (base64 DataURL), inline extraction for supported types.
' - Markdown: conversion via Markdig; HTML emitted with inline CSS/JS (no external assets).
' - Power / resume gating: uses shared flags (powerChanging, IsInResumeCooldown) from other partial.
' - Watchdog interaction: activeRequests increments per request to suppress false restarts.
' - UI integration: selected actions (clipboard, Word insertion, preview) marshalled onto UI thread.
' - Output sanitation: role marker stripping before browser rendering.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Markdig
Imports Microsoft.Office.Interop.Word
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Const InkyBasePath As String = "/inky"
    Private Const InkyUiRoute As String = "/inky"          ' GET → serves HTML
    Private Const InkyApiRoute As String = "/inky/api"     ' POST (JSON) → commands
    Private Const InkyName As String = "Inky"              ' Fallback; AN6 preferred

    Private activeChatId As Integer = 1   ' 1 or 2 – in‑memory only (not persisted)

    Private ReadOnly OriginalSecondaryModelName As String = INI_Model_2

    ''' <summary>
    ''' Represents a single LLM background job (request/response lifecycle, cancellation, file context).
    ''' </summary>
    Private Class LlmJob
        Implements IDisposable
        Public Property Id As String
        Public Property CreatedUtc As DateTime
        Public Property Tcs As TaskCompletionSource(Of String)
        Public Property Cts As CancellationTokenSource
        Public Property UseSecond As Boolean
        Public Property FileObject As String
        Public Sub Dispose() Implements IDisposable.Dispose
            Try
                Cts?.Cancel()
                Cts?.Dispose()
                Tcs?.TrySetCanceled()
            Catch
            End Try
        End Sub
    End Class

    Private ReadOnly jobMap As New System.Collections.Concurrent.ConcurrentDictionary(Of String, LlmJob)()
    Private activeJobs As Integer = 0
    Private Const JobTtlMinutes As Integer = 45

    ''' <summary>
    ''' Handles a single HTTP request (routing: favicon, UI HTML, API POST, CORS preflight).
    ''' Dispatches commands to ProcessRequestInAddIn and writes response bytes.
    ''' </summary>
    Private Async Function HandleHttpRequest(ByVal ctx As System.Net.HttpListenerContext) As System.Threading.Tasks.Task
        Dim req As System.Net.HttpListenerRequest = ctx.Request
        Dim res As System.Net.HttpListenerResponse = ctx.Response

        ' Count in-flight requests
        System.Threading.Interlocked.Increment(activeRequests)

        ' Heartbeat to keep watchdog calm during long processing
        Dim hb As System.Threading.Timer = Nothing

        Try
            hb = New System.Threading.Timer(
                    Sub(stateObj As System.Object)
                        Try
                            lastListenerProgressUtc = System.DateTime.UtcNow
                        Catch
                        End Try
                    End Sub,
                    state:=Nothing,
                    dueTime:=System.TimeSpan.FromSeconds(5),
                    period:=System.TimeSpan.FromSeconds(5))

            If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                Try
                    res = ctx.Response
                    res.StatusCode = 503
                    res.StatusDescription = "Service Unavailable (suspend/resume)"
                    res.AddHeader("Retry-After", "2")
                    res.KeepAlive = False
                    res.Headers("Connection") = "close"
                    res.SendChunked = False
                    Using os = res.OutputStream
                        Dim msgBytes() As System.Byte = System.Text.Encoding.UTF8.GetBytes("Temporarily unavailable during power transition.")
                        res.ContentType = "text/plain; charset=utf-8"
                        res.ContentLength64 = msgBytes.LongLength
                        os.Write(msgBytes, 0, msgBytes.Length)
                    End Using
                    res.Close()
                Catch
                End Try
                Return
            End If

            If IsInResumeCooldown() Then
                Try
                    res = ctx.Response
                    res.StatusCode = 503
                    res.StatusDescription = "Service Unavailable (resume cooldown)"
                    res.AddHeader("Retry-After", "5")
                    res.AddHeader("Access-Control-Allow-Origin", "*")
                    res.KeepAlive = False
                    res.Headers("Connection") = "close"
                    res.SendChunked = False
                    Using os = res.OutputStream
                        Dim msgBytes() As Byte = System.Text.Encoding.UTF8.GetBytes("Resuming from sleep; please retry in a few seconds.")
                        res.ContentType = "text/plain; charset=utf-8"
                        res.ContentLength64 = msgBytes.LongLength
                        os.Write(msgBytes, 0, msgBytes.Length)
                    End Using
                    res.Close()
                Catch
                End Try
                Return
            End If

            ' ---- CORS Preflight ---------------------------------------------------
            If req.HttpMethod.Equals("OPTIONS", System.StringComparison.OrdinalIgnoreCase) Then
                res.AddHeader("Access-Control-Allow-Origin", "*")
                res.AddHeader("Access-Control-Allow-Methods", "GET, POST, PUT, DELETE, OPTIONS")
                res.AddHeader("Access-Control-Allow-Headers", "Content-Type, Authorization")
                res.StatusCode = 204
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.Close()
                Return
            End If

            ' ---- Favicon ----------------------------------------------------------
            If req.HttpMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               req.RawUrl.Equals("/favicon.ico", System.StringComparison.OrdinalIgnoreCase) Then

                Dim png() As System.Byte = GetLogoPngBytes()

                res.ContentType = "image/png"
                res.AddHeader("Cache-Control", "public, max-age=86400")
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = png.LongLength

                Using os As System.IO.Stream = res.OutputStream
                    Await os.WriteAsync(png, 0, png.Length).ConfigureAwait(False)
                End Using
                res.Close()
                Return
            End If

            ' ---- Inky UI (GET /inky[/]) ------------------------------------------
            If req.HttpMethod.Equals("GET", System.StringComparison.OrdinalIgnoreCase) AndAlso
               (req.RawUrl.Equals(InkyUiRoute, System.StringComparison.OrdinalIgnoreCase) OrElse
                req.RawUrl.Equals(InkyUiRoute & "/", System.StringComparison.OrdinalIgnoreCase)) Then

                Dim html As System.String = BuildInkyHtmlPage()
                Dim bufUi() As System.Byte = System.Text.Encoding.UTF8.GetBytes(html)

                res.ContentType = "text/html; charset=utf-8"
                res.AddHeader("Cache-Control", "no-store")
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = bufUi.LongLength

                Using os As System.IO.Stream = res.OutputStream
                    Await os.WriteAsync(bufUi, 0, bufUi.Length).ConfigureAwait(False)
                End Using
                res.Close()
                Return
            End If

            ' ---- Normal flow (POST JSON / API) -----------------------------------
            Dim body As System.String = System.String.Empty
            If req.HasEntityBody Then
                Using rdr As New System.IO.StreamReader(req.InputStream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks:=False, bufferSize:=8192, leaveOpen:=False)
                    body = Await rdr.ReadToEndAsync().ConfigureAwait(False)
                End Using
            End If

            Dim responseText As System.String = Await ProcessRequestInAddIn(body, req.RawUrl).ConfigureAwait(False)
            If responseText Is Nothing Then responseText = System.String.Empty

            ' Content-Type Hints
            Dim contentType As System.String = "text/plain; charset=utf-8"
            If responseText.StartsWith("CT:html" & vbLf, System.StringComparison.Ordinal) Then
                contentType = "text/html; charset=utf-8"
                responseText = responseText.Substring(("CT:html" & vbLf).Length)
            ElseIf responseText.StartsWith("CT:json" & vbLf, System.StringComparison.Ordinal) Then
                contentType = "application/json; charset=utf-8"
                responseText = responseText.Substring(("CT:json" & vbLf).Length)
            End If

            Dim buf() As System.Byte = System.Text.Encoding.UTF8.GetBytes(responseText)

            res.AddHeader("Access-Control-Allow-Origin", "*")
            res.ContentType = contentType
            res.KeepAlive = False
            res.Headers("Connection") = "close"
            res.SendChunked = False
            res.ContentLength64 = buf.LongLength

            Using os As System.IO.Stream = res.OutputStream
                Await os.WriteAsync(buf, 0, buf.Length).ConfigureAwait(False)
            End Using
            res.Close()

        Catch ex As System.Exception
            Try
                Dim err As System.String = "Internal server error: " & ex.Message
                Dim bufErr() As System.Byte = System.Text.Encoding.UTF8.GetBytes(err)

                res.StatusCode = 500
                res.AddHeader("Access-Control-Allow-Origin", "*")
                res.ContentType = "text/plain; charset=utf-8"
                res.KeepAlive = False
                res.Headers("Connection") = "close"
                res.SendChunked = False
                res.ContentLength64 = bufErr.LongLength
                Using os As System.IO.Stream = res.OutputStream
                    os.Write(bufErr, 0, bufErr.Length)
                End Using
                res.Close()
            Catch
            End Try
        Finally
            Try
                If hb IsNot Nothing Then hb.Dispose()
            Catch
            End Try
            System.Threading.Interlocked.Decrement(activeRequests)
            ' Mark progress at the end of a handled request too
            lastListenerProgressUtc = System.DateTime.UtcNow
        End Try
    End Function

    ''' <summary>
    ''' Returns logo PNG bytes; falls back to 1x1 transparent if resource load fails.
    ''' </summary>
    Private Function GetLogoPngBytes() As System.Byte()
        Try
            Using src As System.Drawing.Bitmap = CType(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard).Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Return ms.ToArray()
                End Using
            End Using
        Catch
            ' 1x1 transparent PNG fallback
            Return System.Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQImWNgYGD4DwABdQF+8m3rXQAAAABJRU5ErkJggg==")
        End Try
    End Function


    ' LLM helper (runs off the UI thread) 


    ' 1) Field for scheduler (class/module level)
    Private Shared llmScheduler As System.Threading.Tasks.TaskScheduler


    ' 2) Initialize STA thread with dedicated WinForms message loop


    ''' <summary>
    ''' Ensures a dedicated STA thread with WindowsFormsSynchronizationContext is running for UI-bound operations if required.
    ''' </summary>
    Private Sub EnsureLlmUiThread()
        If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then Return
        If llmScheduler IsNot Nothing Then Return

        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Threading.Tasks.TaskScheduler)()

        llmThread = New System.Threading.Thread(
            Sub()
                System.Threading.SynchronizationContext.SetSynchronizationContext(
                    New System.Windows.Forms.WindowsFormsSynchronizationContext())

                llmSyncContext = System.Threading.SynchronizationContext.Current
                tcs.SetResult(System.Threading.Tasks.TaskScheduler.FromCurrentSynchronizationContext())

                System.Windows.Forms.Application.Run()
            End Sub)

        llmThread.SetApartmentState(System.Threading.ApartmentState.STA)
        llmThread.IsBackground = True
        llmThread.Start()

        llmScheduler = tcs.Task.Result
    End Sub

    ''' <summary>
    ''' Signals STA thread to exit and joins with timeout; clears scheduler and context references.
    ''' </summary>
    Private Sub StopLlmUiThread()
        Try
            If llmSyncContext IsNot Nothing Then
                llmSyncContext.Post(
                    Sub() System.Windows.Forms.Application.ExitThread(),
                    Nothing)
            End If
        Catch
        End Try
        Try
            If llmThread IsNot Nothing AndAlso llmThread.IsAlive Then
                If Not llmThread.Join(2000) Then
                    ' Optional: Log that the thread did not terminate in time.
                End If
            End If
        Catch
        End Try

        llmScheduler = Nothing
        llmSyncContext = Nothing
        llmThread = Nothing
    End Sub

    ''' <summary>
    ''' Runs an LLM operation asynchronously on thread pool with timeout, cancellation propagation, and optional secondary model use.
    ''' Returns model output or status text.
    ''' </summary>
    Public Function RunLlmAsync(
        ByVal sysPrompt As System.String,
        ByVal userPrompt As System.String,
        Optional ByVal UseSecondAPI As System.Boolean = False,
        Optional ByVal ShowTimer As System.Boolean = True,
        Optional ByVal FileObject As System.String = "",
        Optional ByVal cancellationToken As System.Threading.CancellationToken = Nothing
    ) As System.Threading.Tasks.Task(Of System.String)

        ' Use the thread pool – no STA message loop
        Dim effectiveTimeout As Integer = CInt(If(UseSecondAPI, INI_Timeout_2, INI_Timeout))
        ModelTimeout = effectiveTimeout

        Return System.Threading.Tasks.Task.Run(
            Async Function() As System.Threading.Tasks.Task(Of System.String)
                ' Check if we're in a power transition BEFORE starting
                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                    Return "Operation cancelled due to power transition."
                End If

                Using linkedCts As System.Threading.CancellationTokenSource =
                    System.Threading.CancellationTokenSource.CreateLinkedTokenSource(cancellationToken)

                    ' Create a separate timeout CTS that we can control
                    Using timeoutCts As New System.Threading.CancellationTokenSource()
                        ' Add a generous buffer to the model timeout
                        Dim totalTimeout = effectiveTimeout + 60 ' 60 second buffer
                        timeoutCts.CancelAfter(TimeSpan.FromSeconds(totalTimeout))

                        Using combinedCts As System.Threading.CancellationTokenSource =
                            System.Threading.CancellationTokenSource.CreateLinkedTokenSource(linkedCts.Token, timeoutCts.Token)

                            Try
                                Dim llmTask = LLM(sysPrompt, userPrompt, "", "", 0, UseSecondAPI,
                                                  Not ShowTimer, "", FileObject, combinedCts.Token, EnsureUI:=False)

                                ' Wait for either completion or cancellation
                                Dim llmOut As String = Await llmTask.ConfigureAwait(False)

                                ' Mark successful completion
                                lastSuccessfulOperationUtc = System.DateTime.UtcNow

                                If UseSecondAPI AndAlso originalConfigLoaded Then
                                    RestoreDefaults(_context, originalConfig)
                                    originalConfigLoaded = False
                                End If

                                Return If(llmOut, String.Empty)

                            Catch ex As OperationCanceledException When timeoutCts.IsCancellationRequested
                                ' This was a timeout, not user cancellation
                                System.Diagnostics.Debug.WriteLine($"LLM operation timed out after {totalTimeout}s")
                                Return "Operation timed out. Please try again with a shorter prompt or different model."

                            Catch ex As OperationCanceledException
                                ' Check if this is due to power transition
                                If System.Threading.Interlocked.CompareExchange(powerChanging, 0, 0) <> 0 Then
                                    Return "Operation cancelled due to power transition."
                                End If
                                Return "Operation was canceled by the user."

                            Catch ex As System.Exception
                                ' Log the exception for debugging
                                System.Diagnostics.Debug.WriteLine($"LLM Error: {ex.Message}")
                                Return "An error occurred during processing."

                            Finally
                                ' Ensure cleanup happens even during power transitions
                                Try
                                    If UseSecondAPI AndAlso originalConfigLoaded Then
                                        RestoreDefaults(_context, originalConfig)
                                        originalConfigLoaded = False
                                    End If
                                Catch
                                End Try
                            End Try
                        End Using
                    End Using
                End Using
            End Function,
            cancellationToken)
    End Function


    ''' <summary>
    ''' Presents comparison UI and returns user acceptance Boolean (True accepted, False canceled).
    ''' </summary>
    Private Async Function CompareAndInsertSyncConfirm(
        originalText As String,
        llmResult As String
    ) As System.Threading.Tasks.Task(Of Boolean)

        ' Show compare on UI thread
        Await SwitchToUi(Sub() CompareAndInsertText(originalText, llmResult, True)).ConfigureAwait(False)

        ' Await user decision (Esc = False, OK/close = True)
        Dim accepted As Boolean = Await WaitForPreviewDecisionAsync().ConfigureAwait(False)
        Return accepted
    End Function


    ''' <summary>
    ''' Waits for preview dialog dismissal (OK = True, Escape = False). Attaches handlers via polling timer.
    ''' </summary>
    Private Async Function WaitForPreviewDecisionAsync() As System.Threading.Tasks.Task(Of System.Boolean)
        Dim tcs As New System.Threading.Tasks.TaskCompletionSource(Of System.Boolean)()

        ' Handler attachment once on UI thread
        Await SwitchToUi(Sub()

                             Dim previewForm As System.Windows.Forms.Form = Nothing
                             Dim searchTimer As New System.Windows.Forms.Timer() With {.Interval = 100}

                             AddHandler searchTimer.Tick,
                                 Sub()
                                     If previewForm Is Nothing OrElse previewForm.IsDisposed Then
                                         previewForm = System.Windows.Forms.Application.OpenForms _
                                             .Cast(Of System.Windows.Forms.Form)() _
                                             .FirstOrDefault(Function(f As System.Windows.Forms.Form) f.Name = "ShowRTFCustomMessageBox" _
                                                                 OrElse f.Text.StartsWith(AN))

                                         If previewForm Is Nothing Then Return

                                         previewForm.KeyPreview = True

                                         AddHandler previewForm.KeyDown,
                                             Sub(_s As System.Object, e As System.Windows.Forms.KeyEventArgs)
                                                 If e.KeyCode = System.Windows.Forms.Keys.Escape Then
                                                     tcs.TrySetResult(False)
                                                 End If
                                             End Sub

                                         AddHandler previewForm.FormClosed,
                                             Sub(_s As System.Object, _e As System.Windows.Forms.FormClosedEventArgs)
                                                 tcs.TrySetResult(True)
                                                 ' Failsafe: dispose timer
                                                 If searchTimer.Enabled Then
                                                     searchTimer.Stop()
                                                     searchTimer.Dispose()
                                                 End If
                                             End Sub
                                     End If

                                     If tcs.Task.IsCompleted Then
                                         searchTimer.Stop()
                                         searchTimer.Dispose() ' Patch C
                                     End If
                                 End Sub

                             searchTimer.Start()
                         End Sub).ConfigureAwait(False)

        ' IMPORTANT: In async function Task(Of Boolean) → Await tcs.Task
        Return Await tcs.Task.ConfigureAwait(False)
    End Function


    ' Chatbot 


    ''' <summary>
    ''' Attempts to read a setting value by reflection; returns fallback if missing or invalid.
    ''' </summary>
    Private Function TryGetAppSetting(Of T)(ByVal key As System.String, ByVal fallback As T) As T
        Try
            Dim p = GetType(My.MySettings).GetProperty(key, System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
            If p IsNot Nothing Then
                Dim v = DirectCast(p.GetValue(My.Settings, Nothing), Object)
                If v IsNot Nothing Then Return DirectCast(v, T)
            End If
        Catch
        End Try
        Return fallback
    End Function

    ''' <summary>
    ''' Returns bot name based on AN6 setting or default fallback.
    ''' </summary>
    Private Function GetBotName() As System.String
        ' Try: My.Settings("AN6") → else "Inky"
        Dim v As System.String = TryGetAppSetting(Of System.String)("AN6", Nothing)
        If Not System.String.IsNullOrWhiteSpace(v) Then Return v
        Return "Inky"
    End Function

    ''' <summary>
    ''' Returns a data URL (PNG base64) for the logo or empty string on failure.
    ''' </summary>
    Private Function GetLogoDataUrl() As System.String
        Try
            Using src As System.Drawing.Bitmap = CType(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard).Clone(), System.Drawing.Bitmap)
                Using ms As New System.IO.MemoryStream()
                    src.Save(ms, System.Drawing.Imaging.ImageFormat.Png)
                    Dim b64 As System.String = System.Convert.ToBase64String(ms.ToArray())
                    Return "data:image/png;base64," & b64
                End Using
            End Using
        Catch
            Return ""
        End Try
    End Function


    ''' <summary>
    ''' Returns display key for model (prefers ModelDescription, falls back to Model).
    ''' </summary>
    Private Function GetModelDisplayKey(ByVal model As SharedLibrary.SharedLibrary.ModelConfig) As System.String
        If model Is Nothing Then Return ""
        ' Prefer descriptive label:
        If Not System.String.IsNullOrWhiteSpace(model.ModelDescription) Then
            Return model.ModelDescription
        End If
        ' Fallback: internal model name
        If Not System.String.IsNullOrWhiteSpace(model.Model) Then
            Return model.Model
        End If
        Return "Model"
    End Function

    ''' <summary>
    ''' Returns localized greeting based on current UI culture.
    ''' </summary>
    Private Function GetFriendlyGreeting() As System.String
        Dim name As System.String = GetBotName()
        Dim tl As System.String
        Try
            tl = System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        Catch
            tl = "en"
        End Try

        Select Case tl
            Case "de" : Return $"Hallo! Ich bin {name}. Wie kann ich helfen?"
            Case "fr" : Return $"Salut ! Je suis {name}. Comment puis-je aider ?"
            Case "it" : Return $"Ciao! Sono {name}. In cosa posso aiutarti?"
            Case "es" : Return $"¡Hola! Soy {name}. ¿En qué puedo ayudarte?"
            Case Else : Return $"Hi! I’m {name}. How can I help?"
        End Select
    End Function

    Dim botName As System.String = GetBotName()
    Dim brandName As System.String = AN
    Dim logoUrl As System.String = GetLogoDataUrl()
    Dim greet As System.String = GetFriendlyGreeting()

    ''' <summary>
    ''' Simple persisted state container for chat.
    ''' </summary>
    <Serializable>
    Private Class InkyState
        Public History As System.Collections.Generic.List(Of ChatTurn) = New System.Collections.Generic.List(Of ChatTurn)()
        Public SelectedModelKey As System.String = ""
        Public UseSecondApi As System.Boolean = False
        Public LastAssistantText As System.String = ""
        Public DarkMode As System.Boolean = False
        Public SupportsFileUploads As System.Boolean = False
    End Class

    ''' <summary>
    ''' Represents a single conversational turn (user or assistant) with markdown/HTML payload and timestamp.
    ''' </summary>
    <Serializable>
    Private Class ChatTurn
        Public Role As System.String   ' "user" or "assistant"
        Public Markdown As System.String
        Public Html As System.String
        Public Utc As System.DateTime
    End Class

    ''' <summary>
    ''' Returns current UI culture two-letter ISO code or "en" fallback.
    ''' </summary>
    Private Function GetUserLanguageTwoLetter() As System.String
        Try
            Return System.Globalization.CultureInfo.CurrentUICulture.TwoLetterISOLanguageName
        Catch
            Return "en"
        End Try
    End Function

    ''' <summary>
    ''' Determines whether file uploads are supported for current model selection.
    ''' </summary>
    Private Function ComputeSupportsFiles(ByVal useSecond As System.Boolean,
                                          ByVal selectedKey As System.String) As System.Boolean
        Try
            ' Primary API (no alternate selected)
            If Not useSecond Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object)
            End If

            ' Second API, default model
            If System.String.IsNullOrWhiteSpace(selectedKey) Then
                Return Not System.String.IsNullOrWhiteSpace(INI_APICall_Object_2)
            End If

            ' Second API, alternate model -> read APICall_Object
            Dim alts As System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.ModelConfig) = Nothing
            Try
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
            Catch
                alts = Nothing
            End Try
            If alts Is Nothing Then Return False

            Dim sel As SharedLibrary.SharedLibrary.ModelConfig =
                alts.FirstOrDefault(Function(m As SharedLibrary.SharedLibrary.ModelConfig)
                                        If m Is Nothing Then Return False
                                        If Not System.String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                           System.String.Equals(m.ModelDescription, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                            Return True
                                        End If
                                        If Not System.String.IsNullOrWhiteSpace(m.Model) AndAlso
                                           System.String.Equals(m.Model, selectedKey, System.StringComparison.OrdinalIgnoreCase) Then
                                            Return True
                                        End If
                                        Return False
                                    End Function)

            If sel Is Nothing Then Return False

            ' Direct access – fallback via reflection if property missing
            Dim v As System.String = Nothing
            Try
                v = sel.APICall_Object
            Catch
                Try
                    Dim p As System.Reflection.PropertyInfo =
                        GetType(SharedLibrary.SharedLibrary.ModelConfig).GetProperty("APICall_Object",
                            System.Reflection.BindingFlags.Public Or System.Reflection.BindingFlags.Instance)
                    If p IsNot Nothing Then
                        Dim o As System.Object = p.GetValue(sel, Nothing)
                        If o IsNot Nothing Then v = System.Convert.ToString(o, System.Globalization.CultureInfo.InvariantCulture)
                    End If
                Catch
                End Try
            End Try

            Return Not System.String.IsNullOrWhiteSpace(v)
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Loads chat state for given (or active) chat id; initializes defaults if empty.
    ''' </summary>
    Private Function LoadInkyState(Optional chatId As Integer = -1) As InkyState
        If chatId = -1 Then chatId = activeChatId
        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Try
            Dim raw As String = ""
            Try : raw = DirectCast(My.Settings.[GetType]().GetProperty(settingKey).GetValue(My.Settings, Nothing), String) : Catch : raw = "" : End Try
            If String.IsNullOrWhiteSpace(raw) Then
                Dim st As New InkyState()
                ' Default dark mode on first empty chat
                st.DarkMode = True
                Return st
            End If
            Dim stLoaded = Newtonsoft.Json.JsonConvert.DeserializeObject(Of InkyState)(raw)
            If stLoaded Is Nothing Then stLoaded = New InkyState()
            Return stLoaded
        Catch
            Dim st As New InkyState() : st.DarkMode = True
            Return st
        End Try
    End Function

    ''' <summary>
    ''' Persists chat state for given (or active) chat id into application settings.
    ''' </summary>
    Private Sub SaveInkyState(st As InkyState, Optional chatId As Integer = -1)
        If chatId = -1 Then chatId = activeChatId
        Dim settingKey As String = If(chatId = 2, "ChatHistory_Inky2", "ChatHistory_Inky")
        Try
            Dim json = Newtonsoft.Json.JsonConvert.SerializeObject(st)
            Try
                My.Settings.[GetType]().GetProperty(settingKey).SetValue(My.Settings, json, Nothing)
                My.Settings.Save()
            Catch
                ' ignore
            End Try
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Converts markdown to HTML using Markdig advanced pipeline; falls back to HTML-encoded text on error.
    ''' </summary>
    Private Function MarkdownToHtml(ByVal md As System.String) As System.String
        Try
            ' Maximum markdown functionality + soft line breaks as <br/>
            Dim pipeline As Markdig.MarkdownPipeline =
                New Markdig.MarkdownPipelineBuilder().
                    UseAdvancedExtensions().
                    UseSoftlineBreakAsHardlineBreak().
                    UsePipeTables().
                    UseGridTables().
                    UseListExtras().
                    UseFootnotes().
                    UseDefinitionLists().
                    UseAbbreviations().
                    UseAutoLinks().
                    UseTaskLists().
                    UseEmojiAndSmiley().
                    UseMathematics().
                    UseFigures().
                    UseGenericAttributes().
                    Build()

            Return Markdig.Markdown.ToHtml(md, pipeline)
        Catch ex As System.Exception
            ' Fallback: safely encode and preserve line breaks
            Return System.Net.WebUtility.HtmlEncode(md).Replace(vbLf, "<br>")
        End Try
    End Function

    ''' <summary>
    ''' Returns a clipped copy of history limited to the specified character cap (most recent retained).
    ''' </summary>
    Private Function CapHistoryToChars(ByVal st As InkyState, ByVal maxChars As System.Int32) As System.Collections.Generic.List(Of ChatTurn)
        If maxChars <= 0 Then Return New System.Collections.Generic.List(Of ChatTurn)(st.History)
        Dim acc As New System.Text.StringBuilder()
        Dim clipped As New System.Collections.Generic.List(Of ChatTurn)()
        ' iterate from the end (most recent) backwards until cap reached
        For i As System.Int32 = st.History.Count - 1 To 0 Step -1
            Dim turn As ChatTurn = st.History(i)
            Dim piece As System.String = $"[{turn.Role}]{turn.Markdown}" & vbLf
            If acc.Length + piece.Length > maxChars Then Exit For
            acc.Insert(0, piece)
            clipped.Insert(0, turn)
        Next
        Return clipped
    End Function

    ''' <summary>
    ''' Returns user-friendly label for currently selected model (primary or secondary).
    ''' </summary>
    Private Function GetSelectedModelLabel(ByVal useSecond As System.Boolean, ByVal selectedKey As System.String) As System.String
        If Not useSecond Then
            Return If(System.String.IsNullOrWhiteSpace(INI_Model), "Default model", INI_Model)
        End If
        If Not System.String.IsNullOrWhiteSpace(selectedKey) Then
            Return selectedKey
        End If
        Return If(System.String.IsNullOrWhiteSpace(INI_Model_2), "Second API model", INI_Model_2)
    End Function

    ' Builds the entire HTML UI (single file; no external assets)
    ''' <summary>
    ''' Generates full HTML (inline CSS + JS) for the chat UI page.
    ''' </summary>
    Private Function BuildInkyHtmlPage() As System.String
        Dim botName As String = GetBotName()
        Dim brandName As String = If(Not String.IsNullOrWhiteSpace(AN), AN, botName)
        Dim logoUrl As String = GetLogoDataUrl()
        Dim greet As String = GetFriendlyGreeting()

        Dim html As New System.Text.StringBuilder()

        html.AppendLine("<!doctype html>")
        html.AppendLine("<html lang=""en""><head><meta charset=""utf-8"">")
        html.AppendLine("<meta name=""viewport"" content=""width=device-width, initial-scale=1"">")
        html.AppendLine("<link rel=""shortcut icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<link rel=""icon"" type=""image/png"" href=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """>")
        html.AppendLine("<title>" & System.Net.WebUtility.HtmlEncode(brandName) & " — Local Chat</title>")

        ' CSS
        html.AppendLine("<style>")
        html.AppendLine(":root{--bg:#0b0f14;--card:#11161d;--fg:#e8eef6;--muted:#9aa8b7;--border:#1b2430;--border-strong:#2d3744;--elev:#1a222c;--press-shadow:inset 0 2px 6px rgba(0,0,0,.45);}")
        html.AppendLine(":root.light{--bg:#f6f7f9;--card:#ffffff;--fg:#0e1116;--muted:#5d6a77;--border:#e2e5e9;--border-strong:#c9cfd6;--elev:#eef1f4;--press-shadow:inset 0 2px 5px rgba(0,0,0,.08);}")
        html.AppendLine("html,body{height:100%;margin:0;font-family:system-ui,Segoe UI,Roboto,Arial,sans-serif;background:var(--bg);color:var(--fg);}")
        html.AppendLine(".wrap{display:flex;flex-direction:column;height:100%;}")
        html.AppendLine(".topbar{display:flex;gap:.5rem;align-items:center;padding:.75rem 1rem;border-bottom:1px solid var(--border);background:var(--card);position:sticky;top:0;z-index:5;flex-wrap:nowrap;overflow:hidden;}")
        html.AppendLine(".topline{display:flex;align-items:center;gap:.6rem;min-width:0;}")
        html.AppendLine(".topline img.logo{width:24px;height:24px;border-radius:6px;display:block}")
        html.AppendLine(".topline .brandbig{font-weight:700;white-space:nowrap;}")
        html.AppendLine(".topline .sub{color:var(--muted);font-size:.9rem;white-space:nowrap;}")
        html.AppendLine(".muted{color:var(--muted);font-size:.85rem}")
        html.AppendLine(".spacer{flex:0 0 0;}")
        html.AppendLine("select,button,input,textarea{background:var(--card);color:var(--fg);border:1px solid var(--border);border-radius:.6rem;font:inherit;}")
        html.AppendLine("select,button,input{padding:.5rem .7rem;}")
        html.AppendLine("button{cursor:pointer;transition:background .16s,filter .12s,transform .08s,box-shadow .18s;}")
        html.AppendLine("button:hover{filter:brightness(1.07)}")
        html.AppendLine("button:disabled{opacity:.5;cursor:not-allowed}")
        html.AppendLine("button.is-pressed, .chatTab.is-pressed{transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.92);}")
        html.AppendLine("button:active:not(:disabled){transform:translateY(1px);box-shadow:var(--press-shadow);filter:brightness(.9);}")
        html.AppendLine(".chat{flex:1;overflow:auto;padding:1rem;}")
        html.AppendLine(".row{display:flex;margin:0 auto 1rem auto;max-width:1000px;padding:0 .25rem;}")
        html.AppendLine(".row.bot{justify-content:flex-start}")
        html.AppendLine(".row.user{justify-content:flex-end}")
        html.AppendLine(".bubble{max-width:75%;padding:1rem;border:1px solid var(--border);background:var(--card);border-radius:1rem;box-shadow:0 1px 3px rgba(0,0,0,.25)}")
        html.AppendLine(".bot .bubble{border-top-right-radius:.35rem}")
        html.AppendLine(".user .bubble{border-top-left-radius:.35rem}")
        html.AppendLine(".role{font-size:.75rem;color:var(--muted);margin-bottom:.25rem}")
        html.AppendLine(".inputbar{display:flex;gap:.5rem;padding:1rem;border-top:1px solid var(--border);background:var(--card)}")
        html.AppendLine("textarea{flex:1;resize:vertical;min-height:52px;max-height:220px;border-radius:.8rem;padding:.75rem;line-height:1.25;}")
        html.AppendLine(".hint{font-size:.7rem;letter-spacing:.3px;color:var(--muted);padding:.25rem 1rem 1rem}")
        html.AppendLine("a{color:inherit;text-decoration:underline;text-decoration-color:rgba(255,255,255,.35)}")
        html.AppendLine(":root.light a{text-decoration-color:rgba(0,0,0,.4)}")
        html.AppendLine("a:hover{filter:brightness(1.15)}")
        html.AppendLine("code,pre{font-family:ui-monospace,Consolas,monospace;font-size:.85rem}")
        html.AppendLine("pre{overflow:auto;padding:.75rem;border:1px solid var(--border);border-radius:.6rem;position:relative;background:var(--elev);}")
        html.AppendLine(".code-copy-btn{position:absolute;top:6px;right:6px;padding:4px 8px;font-size:.65rem;line-height:1;border:1px solid var(--border);border-radius:4px;background:rgba(0,0,0,.45);backdrop-filter:blur(3px);cursor:pointer;display:flex;align-items:center;gap:6px;color:var(--fg);opacity:0;transition:opacity .18s,background .18s;}")
        html.AppendLine("pre:hover .code-copy-btn{opacity:1}")
        html.AppendLine(".code-copy-btn svg{width:16px;height:16px;display:block}")
        html.AppendLine(".code-copy-btn.copied{background:#2c3440;color:#fff}")
        html.AppendLine(":root.light .code-copy-btn.copied{background:#d5d9dd;color:#111}")
        html.AppendLine(".code-copy-btn:focus{outline:2px solid var(--border-strong);}")
        html.AppendLine(".chatTab{padding:.45rem .55rem;min-width:32px;font-size:.7rem;font-weight:600;line-height:1;border:1px solid var(--border);background:var(--card);color:var(--muted);transition:background .18s,border-color .18s,color .18s,transform .08s,box-shadow .18s;flex-shrink:0;}")
        html.AppendLine(".chatTab:hover:not(:disabled){background:var(--elev);color:var(--fg);}")
        html.AppendLine(".chatTab.active{background:#222b35;border-color:var(--border-strong);color:#fff;box-shadow:inset 0 0 0 1px #303c46;}")
        html.AppendLine(":root.light .chatTab.active{background:#e2e5e9;border-color:var(--border-strong);color:#0e1116;box-shadow:inset 0 0 0 1px #c9cfd6;}")
        html.AppendLine(".chatTab:focus{outline:2px solid var(--border-strong);outline-offset:1px;}")
        ' Typing + elapsed
        html.AppendLine(".typing-dots{display:inline-flex;gap:6px;align-items:center;}")
        html.AppendLine(".typing-dots span{width:7px;height:7px;border-radius:50%;background:currentColor;opacity:.35;animation:tdots 1.2s infinite ease-in-out;}")
        html.AppendLine(".typing-dots span:nth-child(2){animation-delay:.2s}")
        html.AppendLine(".typing-dots span:nth-child(3){animation-delay:.4s}")
        html.AppendLine("@keyframes tdots{0%,80%,100%{transform:translateY(0);opacity:.3}40%{transform:translateY(-5px);opacity:.85}}")
        html.AppendLine(".typing-elapsed{margin-left:8px;font-size:.65rem;color:var(--muted);font-family:ui-monospace,monospace;opacity:.8;}")
        html.AppendLine(".actions{display:flex;flex-direction:row;gap:.5rem;align-items:stretch;}")
        html.AppendLine(".actions .stack{display:flex;flex-direction:column;gap:.5rem;}")
        html.AppendLine("#cancelBtn{display:none;align-self:stretch;height:auto;}")
        ' Responsive model select behaviour
        html.AppendLine("#modelSel{flex:1 1 260px;max-width:420px;min-width:110px;overflow:hidden;white-space:nowrap;text-overflow:ellipsis;box-sizing:border-box;}")
        html.AppendLine("#modelSel.squeezed{max-width:55vw;}")
        html.AppendLine(".topbar button,#modelSel{flex-shrink:0;}")
        html.AppendLine("@media (max-width:1000px){#modelSel{max-width:360px;}}")
        html.AppendLine("@media (max-width:880px){.topline .sub{display:none;}}")
        html.AppendLine("@media (max-width:760px){.topline .brandbig{max-width:140px;overflow:hidden;text-overflow:ellipsis;}#modelSel{max-width:300px;}}")
        html.AppendLine("@media (max-width:640px){#modelSel{max-width:55vw;} .topline .brandbig{max-width:110px;}}")
        html.AppendLine("@media (max-width:640px){.actions{flex-direction:column;}.actions .stack{flex-direction:row;}.actions .stack button{flex:1;}#cancelBtn{align-self:auto;height:auto;}}")
        html.AppendLine("</style>")

        html.AppendLine("</head><body>")
        html.AppendLine("<div class=""wrap"">")

        ' Top bar
        html.AppendLine("  <div class=""topbar"">")
        html.AppendLine("    <div class=""topline"">")
        If Not String.IsNullOrWhiteSpace(logoUrl) Then
            html.AppendLine("      <img class=""logo"" src=""" & System.Net.WebUtility.HtmlEncode(logoUrl) & """ alt=""logo"">")
        End If
        html.AppendLine("      <div class=""brandbig"">" & System.Net.WebUtility.HtmlEncode(brandName) & "</div>")
        html.AppendLine("      <div class=""sub"">Local Chat</div>")
        html.AppendLine("    </div>")
        html.AppendLine("    <div class=""spacer""></div>")
        html.AppendLine("    <select id=""modelSel"" title=""Model""></select>")
        html.AppendLine("    <button id=""copyBtn"" title=""Copy last answer to clipboard"">Copy last</button>")
        html.AppendLine("    <button id=""toWordBtn"" title=""Move this chat thread into a new Word document"">To Word</button>")
        html.AppendLine("    <button id=""clearBtn"" title=""Clear current conversation"">Clear</button>")
        html.AppendLine("    <button id=""chat1Btn"" class=""chatTab"" data-chat=""1"" title=""Chat 1"">1</button>")
        html.AppendLine("    <button id=""chat2Btn"" class=""chatTab"" data-chat=""2"" title=""Chat 2"">2</button>")
        html.AppendLine("    <button id=""themeBtn"" title=""Toggle theme"">Theme</button>")
        html.AppendLine("  </div>")

        html.AppendLine("  <div id=""chat"" class=""chat""></div>")

        html.AppendLine("  <div class=""inputbar"">")
        html.AppendLine("    <textarea id=""msg"" placeholder=""" & System.Net.WebUtility.HtmlEncode(greet) & """ autofocus></textarea>")
        html.AppendLine("    <div class=""actions"">" &
                    "<div class=""stack"">" &
                        "<button id=""sendBtn"">Send</button>" &
                        "<button id=""pureBtn"" title=""Send only this raw text (no system prompt, no history)"">Pure</button>" &
                    "</div>" &
                    "<button id=""cancelBtn"" style=""display:none;"">Cancel</button>" &
                "</div>")
        html.AppendLine("  </div>")
        html.AppendLine("  <div class=""hint"">Drag & drop a file (only visible to the chatbot for the current prompt) • Enter=send • Shift+Enter=newline • Ctrl+L=clear</div>")
        html.AppendLine("</div>")

        ' JS
        html.AppendLine("<script>")
        html.AppendLine("window.__botName=" & Newtonsoft.Json.JsonConvert.SerializeObject(botName) & ";")
        html.AppendLine("let __supportsFiles=false;")
        html.AppendLine("let __pendingFilePath='';")
        html.AppendLine("let dark=false;")
        html.AppendLine("let __currentJobId=null;")
        html.AppendLine("let __jobCanceled=false;")
        html.AppendLine("let __typingBubbleId=null;")
        html.AppendLine("let __jobStartTs=0;")
        html.AppendLine("let __elapsedTimer=null;")

        ' Press feedback
        html.AppendLine("(function(){const pressOn=e=>{const b=e.target.closest('button');if(!b||b.disabled)return;b.classList.add('is-pressed');};const pressOff=()=>{document.querySelectorAll('button.is-pressed').forEach(b=>b.classList.remove('is-pressed'));};['mousedown','touchstart'].forEach(ev=>document.addEventListener(ev,pressOn,{passive:true}));['mouseup','mouseleave','blur'].forEach(ev=>document.addEventListener(ev,pressOff));document.addEventListener('keydown',e=>{if((e.key===' '||e.key==='Enter')){const b=e.target.closest('button');if(b&&!b.disabled)b.classList.add('is-pressed');}});document.addEventListener('keyup',e=>{if(e.key===' '||e.key==='Enter')pressOff();});})();")

        ' Helpers
        html.AppendLine("function copyText(t){if(navigator.clipboard){return navigator.clipboard.writeText(t);}return new Promise((res,rej)=>{try{const ta=document.createElement('textarea');ta.value=t;ta.style.position='fixed';ta.style.left='-9999px';document.body.appendChild(ta);ta.select();document.execCommand('copy');ta.remove();res();}catch(e){rej(e);}});}")
        html.AppendLine("function enhanceCodeBlocks(scope){(scope||document).querySelectorAll('pre').forEach(pre=>{if(pre.dataset.enhanced==='1')return;const btn=document.createElement('button');btn.type='button';btn.className='code-copy-btn';btn.innerHTML='<svg viewBox=""0 0 24 24"" fill=""none"" stroke=""currentColor"" stroke-width=""2"" stroke-linecap=""round"" stroke-linejoin=""round""><rect x=""9"" y=""9"" width=""13"" height=""13"" rx=""2"" ry=""2""/><path d=""M5 15H4a2 2 0 0 1-2-2V4c0-1.1.9-2 2-2h9a2 2 0 0 1 2 2v1""/></svg>';btn.addEventListener('click',()=>{const code=pre.querySelector('code');const txt=code?code.innerText:pre.innerText;copyText(txt).then(()=>{btn.classList.add('copied');setTimeout(()=>btn.classList.remove('copied'),1500);});});pre.appendChild(btn);pre.dataset.enhanced='1';});}")
        html.AppendLine("const api=async(cmd,data={})=>{try{const r=await fetch('/inky/api',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(Object.assign({Command:cmd},data))});const txt=await r.text();try{return JSON.parse(txt);}catch{return{ok:false,error:txt}}}catch(e){return{ok:false,error:e.message||'Network error'}}};")

        html.AppendLine("const chatEl=document.getElementById('chat');")
        html.AppendLine("const msgEl=document.getElementById('msg');")
        html.AppendLine("const modelSel=document.getElementById('modelSel');")
        html.AppendLine("const copyBtn=document.getElementById('copyBtn');")
        html.AppendLine("const toWordBtn=document.getElementById('toWordBtn');")
        html.AppendLine("const clearBtn=document.getElementById('clearBtn');")
        html.AppendLine("const themeBtn=document.getElementById('themeBtn');")
        html.AppendLine("const cancelBtn=document.getElementById('cancelBtn');")
        html.AppendLine("const chat1Btn=document.getElementById('chat1Btn');")
        html.AppendLine("const chat2Btn=document.getElementById('chat2Btn');")
        html.AppendLine("const sendBtn=document.getElementById('sendBtn');")
        html.AppendLine("const pureBtn=document.getElementById('pureBtn');")
        html.AppendLine("const topbar=document.querySelector('.topbar');")

        html.AppendLine("function setTheme(isDark){dark=!!isDark;document.documentElement.classList.toggle('light',!dark);} ")
        html.AppendLine("function forceExternalLinks(scope){try{(scope||document).querySelectorAll('a[href]').forEach(a=>{a.target='_blank';a.rel='noopener noreferrer';});}catch{}}")
        html.AppendLine("function setActiveChatBtn(id){document.querySelectorAll('.chatTab').forEach(b=>b.classList.toggle('active',b.dataset.chat==String(id)));}")
        html.AppendLine("function disableChatSwitch(dis){chat1Btn.disabled=dis;chat2Btn.disabled=dis;}")

        ' Tooltip updates & responsive adjustment
        html.AppendLine("function updateModelTooltip(){try{if(!modelSel) return;const opt=modelSel.options[modelSel.selectedIndex];if(opt){modelSel.title=opt.textContent||'Model';}}catch{}}")
        html.AppendLine("function adjustModelSel(){if(!topbar) return;requestAnimationFrame(()=>{if(topbar.scrollWidth>topbar.clientWidth){modelSel.classList.add('squeezed');}else{modelSel.classList.remove('squeezed');}});}")
        html.AppendLine("window.addEventListener('resize',adjustModelSel);")

        html.AppendLine("function render(turns){chatEl.innerHTML='';for(const t of (turns||[])){const row=document.createElement('div');row.className='row '+(t.role==='user'?'user':'bot');const bub=document.createElement('div');bub.className='bubble';const rl=document.createElement('div');rl.className='role';rl.textContent=(t.role==='user'?'You':(window.__botName||'Bot'));bub.appendChild(rl);const cont=document.createElement('div');if(t && t.html){cont.innerHTML=t.html;forceExternalLinks(cont);}else if(t && t.markdown){const safe=t.markdown.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>');cont.innerHTML=safe;}bub.appendChild(cont);row.appendChild(bub);chatEl.appendChild(row);}chatEl.scrollTop=chatEl.scrollHeight;enhanceCodeBlocks(chatEl);} ")
        html.AppendLine("function addTempAssistantBubble(html){const id='tmp-'+Math.random().toString(36).slice(2);chatEl.insertAdjacentHTML('beforeend',`<div class=""row bot"" id=""${id}""><div class=""bubble""><div class=""role"">${window.__botName||'Bot'}</div><div class=""tmpContent"">${html}</div></div></div>`);chatEl.scrollTop=chatEl.scrollHeight;return id;}")
        html.AppendLine("function removeTempBubble(id){const el=document.getElementById(id);if(el)el.remove();}")
        html.AppendLine("function replaceAssistantBubble(id,html){const row=document.getElementById(id);if(!row)return;const c=row.querySelector('.tmpContent');if(c){c.innerHTML=html;forceExternalLinks(row);enhanceCodeBlocks(row);}}")

        ' Typing + elapsed
        html.AppendLine("function ensureTypingBubble(){if(__typingBubbleId)return;const content='<div class=""typing-container""><span class=""typing-dots""><span></span><span></span><span></span></span><span id=""typingElapsed"" class=""typing-elapsed"" style=""display:none;"">(0s)</span></div>';__typingBubbleId=addTempAssistantBubble(content);}")
        html.AppendLine("function updateElapsed(){if(!__typingBubbleId)return;const el=document.getElementById('typingElapsed');if(!el)return;const sec=Math.floor((Date.now()-__jobStartTs)/1000);if(sec>=10){el.style.display='inline-block';el.textContent='(' + sec + 's)';}}")
        html.AppendLine("function startElapsedTimer(){stopElapsedTimer();__jobStartTs=Date.now();__elapsedTimer=setInterval(updateElapsed,1000);}")
        html.AppendLine("function stopElapsedTimer(){if(__elapsedTimer){clearInterval(__elapsedTimer);__elapsedTimer=null;}const el=document.getElementById('typingElapsed');if(el)el.style.display='none';}")
        html.AppendLine("function removeTypingBubble(){if(__typingBubbleId){removeTempBubble(__typingBubbleId);__typingBubbleId=null;}stopElapsedTimer();}")

        ' boot
        html.AppendLine("async function boot(){const st=await api('inky_getstate');if(!st.ok){alert(st.error||'Init failed');return;}__supportsFiles=(st.supportsFiles===true);setTheme(st.darkMode!==false);render(st.history||[]);modelSel.innerHTML='';for(const m of (st.models||[])){const o=document.createElement('option');o.value=m.key||'';o.textContent=m.label||'';o.disabled=!!m.disabled;o.title=o.textContent;if(m.selected&&!o.disabled)o.selected=true;modelSel.appendChild(o);}if(!modelSel.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}updateModelTooltip();if(st.greeting && (!Array.isArray(st.history)||st.history.length===0)){msgEl.placeholder=st.greeting;}setActiveChatBtn(st.activeChat||1);adjustModelSel();} ")

        ' pollJob
        html.AppendLine("async function pollJob(jobId){if(!jobId)return;__currentJobId=jobId;__jobCanceled=false;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);try{for(;;){await new Promise(r=>setTimeout(r,2000));if(__jobCanceled)break;const s=await api('inky_jobstatus',{Job:jobId});if(!s.ok){console.warn('job status error',s.error);break;}if(s.status==='running'){continue;}const st=await api('inky_getstate');if(st.ok){render(st.history||[]);}break;} }finally{cancelBtn.style.display='none';removeTypingBubble();sendBtn.disabled=false;pureBtn.disabled=false;disableChatSwitch(false);__currentJobId=null;adjustModelSel();}}")

        ' send (normal)
        html.AppendLine("async function send(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${t.replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_send',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' pureSend
        html.AppendLine("async function pureSend(){if(__currentJobId){return;}const t=msgEl.value.trim();if(!t)return;msgEl.value='';sendBtn.disabled=true;pureBtn.disabled=true;chatEl.insertAdjacentHTML('beforeend',`<div class=""row user""><div class=""bubble""><div class=""role"">You</div><div>${('Pure: '+t).replaceAll('&','&amp;').replaceAll('<','&lt;').replaceAll('>','&gt;').replaceAll('\n','<br>')}</div></div></div>`);let typingId=addTempAssistantBubble('<span class=""typing-dots""><span></span><span></span><span></span></span>');const payload={Text:t};if(__pendingFilePath)payload.FileObject=__pendingFilePath;let r;try{r=await api('inky_pure',payload);}catch(e){r={ok:false,error:e.message||'Network error'};}if(!r||!r.ok){removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;alert(r&&r.error||'Error');__pendingFilePath='';adjustModelSel();return;}__pendingFilePath='';if(r.job){if(r.history){render(r.history||[]);}removeTempBubble(typingId);__typingBubbleId=null;ensureTypingBubble();startElapsedTimer();cancelBtn.style.display='inline-block';disableChatSwitch(true);pollJob(r.job);}else{removeTempBubble(typingId);sendBtn.disabled=false;pureBtn.disabled=false;if(r.history){render(r.history||[]);}adjustModelSel();}}")

        ' drag/drop
        html.AppendLine("(function(){const stop=e=>{e.preventDefault();e.stopPropagation();};['dragenter','dragover','dragleave','drop'].forEach(ev=>document.addEventListener(ev,stop,false));document.addEventListener('drop',async e=>{const files=[...(e.dataTransfer&&e.dataTransfer.files)||[]];if(!files.length)return;const f=files[0];if(!__supportsFiles){addTempAssistantBubble('File uploads are not supported for the current model.');return;}const tempId=addTempAssistantBubble(`Uploading <b>${f.name.replaceAll('&','&amp;')}</b> (${(f.size/1024).toFixed(1)} KB)…`);try{const fr=new FileReader();const dataUrl=await new Promise((res,rej)=>{fr.onerror=()=>rej(new Error('read error'));fr.onload=()=>res(fr.result);fr.readAsDataURL(f);});const r=await api('inky_upload',{Name:f.name,DataUrl:String(dataUrl||'')});if(!r.ok){replaceAssistantBubble(tempId,'Upload failed: '+(r.error||'unknown'));return;}if(r.supported===false){replaceAssistantBubble(tempId,'File uploads are not supported for this model.');return;}__pendingFilePath=r.path||'';replaceAssistantBubble(tempId,`Added file: <b>${(r.name||f.name).replaceAll('&','&amp;')}</b>`);}catch(err){replaceAssistantBubble(tempId,'Upload failed: '+(err&&err.message?err.message:'unknown'));}} ,false);})();")

        ' events
        html.AppendLine("modelSel.addEventListener('change',async()=>{if(__currentJobId)return;const opt=modelSel.options[modelSel.selectedIndex];if(!opt||opt.disabled||!opt.value){const fe=[...modelSel.options].find(o=>!o.disabled&&o.value);if(fe)fe.selected=true;}const r=await api('inky_setmodel',{Key:opt.value});updateModelTooltip();adjustModelSel();if(!r.ok){alert(r.error||'Failed to set model');return;}if(typeof r.supportsFiles==='boolean')__supportsFiles=r.supportsFiles;});")
        html.AppendLine("clearBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_clear');if(r.ok){render([]);if(r.greeting)msgEl.placeholder=r.greeting;}else{alert(r.error||'Failed to clear');}adjustModelSel();});")
        html.AppendLine("copyBtn.addEventListener('click',async()=>{const r=await api('inky_copylast');if(!r.ok){alert(r.error||'Nothing to copy')}});")
        html.AppendLine("toWordBtn.addEventListener('click',async()=>{if(__currentJobId)return;const r=await api('inky_toword');if(!r.ok){alert(r.error||'Failed to create Word document')}});")
        html.AppendLine("themeBtn.addEventListener('click',async()=>{if(__currentJobId)return;const target=!dark;setTheme(target);const r=await api('inky_toggletheme');if(!r.ok){setTheme(!target);alert(r.error||'Theme switch failed');return;}if(typeof r.darkMode==='boolean')setTheme(r.darkMode===true);adjustModelSel();});")
        html.AppendLine("msgEl.addEventListener('keydown',e=>{if(e.key==='Enter'&&!e.shiftKey){e.preventDefault();send();}if(e.ctrlKey&&e.key.toLowerCase()==='l'){e.preventDefault();clearBtn.click();}});")
        html.AppendLine("sendBtn.addEventListener('click',send);")
        html.AppendLine("pureBtn.addEventListener('click',pureSend);")
        html.AppendLine("cancelBtn.addEventListener('click',async()=>{if(!__currentJobId)return;__jobCanceled=true;await api('inky_cancel',{Job:__currentJobId});});")
        html.AppendLine("chatEl.addEventListener('click',e=>{const a=e.target&&e.target.closest&&e.target.closest('a[href]');if(!a)return;if(a.target!=='_blank'){a.target='_blank';a.rel='noopener noreferrer';}});")
        html.AppendLine("async function switchChat(n){if(__currentJobId)return;const r=await api('inky_switch',{Chat:String(n)});if(!r.ok){alert(r.error||'Switch failed');return;}setActiveChatBtn(r.activeChat||n);render(r.history||[]);if(r.greeting){msgEl.placeholder=r.greeting;}updateModelTooltip();adjustModelSel();}")
        html.AppendLine("chat1Btn.addEventListener('click',()=>switchChat(1));")
        html.AppendLine("chat2Btn.addEventListener('click',()=>switchChat(2));")

        html.AppendLine("boot();")
        html.AppendLine("</script>")
        html.AppendLine("</body></html>")
        Return html.ToString()
    End Function

    ''' <summary>
    ''' Wraps object into JSON response with ok=True plus optional fields.
    ''' </summary>
    Private Function JsonOk(o As Object) As System.String
        Return "CT:json" & vbLf & Newtonsoft.Json.JsonConvert.SerializeObject(o)
    End Function

    ''' <summary>
    ''' Wraps error message into JSON response with ok=False.
    ''' </summary>
    Private Function JsonErr(msg As System.String) As System.String
        Return JsonOk(New With {.ok = False, .error = msg})
    End Function

    ' ===== (D) EXTEND ProcessRequestInAddIn with the Inky commands =====

    Private Shared ReadOnly AlternateModelLock As New Object

    ''' <summary>
    ''' Central dispatcher for browser API commands under InkyApiRoute; handles chat operations,
    ''' job control, model selection, theme toggle, file upload, and falls back to legacy commands.
    ''' </summary>
    Private Async Function ProcessRequestInAddIn(
        body As System.String,
        rawUrl As System.String) As System.Threading.Tasks.Task(Of System.String)

        ' If this is a browser POST to our Inky API, j may be JSON; otherwise keep your existing flow
        If rawUrl IsNot Nothing AndAlso rawUrl.StartsWith(InkyApiRoute, System.StringComparison.OrdinalIgnoreCase) Then
            Try
                Dim j As Newtonsoft.Json.Linq.JObject = If(
                    Not System.String.IsNullOrWhiteSpace(body),
                    Newtonsoft.Json.Linq.JObject.Parse(body),
                    New Newtonsoft.Json.Linq.JObject())
                Dim cmd As System.String = j("Command")?.ToString()

                Select Case cmd
                    Case "inky_getstate"
                        Dim st As InkyState = LoadInkyState()

                        Try
                            st.DarkMode = My.Settings.Inky_DarkMode
                        Catch
                        End Try

                        ' Re-compute per current selection on every getstate
                        Try
                            st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                            SaveInkyState(st)
                        Catch
                            st.SupportsFileUploads = False
                        End Try

                        Dim greeting As System.String = Nothing
                        If st.History.Count = 0 Then greeting = GetFriendlyGreeting()

                        Dim models As System.Collections.Generic.List(Of System.Object) =
                            Await GetModelListForBrowserAsync(st)

                        Return JsonOk(New With {
                            .ok = True,
                            .history = ToBrowserTurns(LoadInkyState().History),
                            .greeting = greeting,
                            .models = models,
                            .modelLabel = GetSelectedModelLabel(st.UseSecondApi, st.SelectedModelKey),
                            .darkMode = st.DarkMode,
                            .supportsFiles = st.SupportsFileUploads,
                            .activeChat = activeChatId
                        })

                    ' Remaining cases kept intact (command-specific logic)
                    ' ------------------------------------------------------------------
                    Case "inky_switch"
                        ' (Switch chat logic preserved)
                        ' ------------------------------------------------------------------
                        Dim which As String = j("Chat")?.ToString()
                        activeChatId = If(which = "2", 2, 1)
                        Dim stSw = LoadInkyState()
                        If stSw.History.Count = 0 AndAlso Not stSw.DarkMode Then stSw.DarkMode = True
                        Dim greetingSwitch As String = If(stSw.History.Count = 0, GetFriendlyGreeting(), Nothing)

                        Try
                            My.Settings.Inky_LastChat = activeChatId
                            My.Settings.Save()
                        Catch
                        End Try

                        Return JsonOk(New With {
                            .ok = True,
                            .history = ToBrowserTurns(stSw.History),
                            .activeChat = activeChatId,
                            .darkMode = stSw.DarkMode,
                            .supportsFiles = ComputeSupportsFiles(stSw.UseSecondApi, stSw.SelectedModelKey),
                            .greeting = greetingSwitch
                        })

                    Case "inky_upload"
                        ' (Upload logic preserved)
                        ' ------------------------------------------------------------------
                        Try
                            Dim stU As InkyState = LoadInkyState()
                            Dim supports As System.Boolean = False
                            Try
                                supports = ComputeSupportsFiles(stU.UseSecondApi, stU.SelectedModelKey)
                            Catch
                                supports = False
                            End Try
                            If Not supports Then
                                Return JsonOk(New With {.ok = True, .supported = False})
                            End If
                            Dim name As System.String = j("Name")?.ToString()
                            Dim dataUrl As System.String = j("DataUrl")?.ToString()
                            If System.String.IsNullOrWhiteSpace(name) OrElse System.String.IsNullOrWhiteSpace(dataUrl) Then
                                Return JsonErr("Missing file data.")
                            End If
                            Dim commaIx As System.Int32 = dataUrl.IndexOf(","c)
                            If commaIx < 0 Then Return JsonErr("Bad DataURL.")
                            Dim b64 As System.String = dataUrl.Substring(commaIx + 1)
                            Dim bytes() As System.Byte
                            Try
                                bytes = System.Convert.FromBase64String(b64)
                            Catch exB64 As System.Exception
                                Return JsonErr("Invalid base64: " & exB64.Message)
                            End Try
                            Dim dir As System.String = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "InkyUploads")
                            If Not System.IO.Directory.Exists(dir) Then System.IO.Directory.CreateDirectory(dir)
                            Dim safeName As System.String = System.IO.Path.GetFileName(name)
                            For Each c As System.Char In System.IO.Path.GetInvalidFileNameChars()
                                safeName = safeName.Replace(c, "_"c)
                            Next
                            Dim unique As System.String = System.Guid.NewGuid().ToString("N")
                            Dim target As System.String = System.IO.Path.Combine(dir, unique & "_" & safeName)
                            System.IO.File.WriteAllBytes(target, bytes)
                            Return JsonOk(New With {.ok = True, .supported = True, .path = target, .name = safeName, .size = bytes.LongLength})
                        Catch exUp As System.Exception
                            Return JsonErr("Upload failed: " & exUp.Message)
                        End Try

                    Case "inky_cancel"
                        ' (Cancellation logic preserved)
                        ' ------------------------------------------------------------------
                        ' Optional job id (preferred modern path)
                        Dim jobId As String = j("Job")?.ToString()

                        If Not String.IsNullOrWhiteSpace(jobId) Then
                            Dim job As LlmJob = Nothing
                            If Not jobMap.TryGetValue(jobId, job) Then
                                Return JsonErr("Unknown job.")
                            End If
                            Try
                                If job Is Nothing OrElse job.Cts Is Nothing Then
                                    Return JsonErr("Job has no cancellation handle.")
                                End If
                                If Not job.Cts.IsCancellationRequested Then
                                    job.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "cancelRequested"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = jobId,
                                        .status = "alreadyCanceled"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Cancel failed: " & ex.Message)
                            End Try
                        End If
                        ' Legacy path
                        If llmOperationCts IsNot Nothing AndAlso Not llmOperationCts.IsCancellationRequested Then
                            Try
                                llmOperationCts.Cancel()
                                Return JsonOk(New With {
                                    .ok = True,
                                    .status = "cancelRequestedLegacy"
                                })
                            Catch ex As System.Exception
                                Return JsonErr("Legacy cancel failed: " & ex.Message)
                            End Try
                        End If
                        ' Fallback
                        Dim fallbackJob As LlmJob =
                            jobMap.Values _
                                  .Where(Function(x) x IsNot Nothing AndAlso
                                                     x.Tcs IsNot Nothing AndAlso
                                                     Not x.Tcs.Task.IsCompleted) _
                                  .OrderByDescending(Function(x) x.CreatedUtc) _
                                  .FirstOrDefault()
                        If fallbackJob IsNot Nothing Then
                            Try
                                If Not fallbackJob.Cts.IsCancellationRequested Then
                                    fallbackJob.Cts.Cancel()
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "cancelRequestedFallback"
                                    })
                                Else
                                    Return JsonOk(New With {
                                        .ok = True,
                                        .job = fallbackJob.Id,
                                        .status = "alreadyCanceledFallback"
                                    })
                                End If
                            Catch ex As System.Exception
                                Return JsonErr("Fallback cancel failed: " & ex.Message)
                            End Try
                        End If
                        Return JsonErr("No active operation to cancel.")

                    Case "inky_send"
                        ' ------------------ (A) Read request & validate ------------------
                        Dim fileObject As System.String = j("FileObject")?.ToString()
                        Dim uploadedTempPath As System.String = fileObject
                        Dim textBody As System.String = j("Text")?.ToString()
                        If System.String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If
                        Dim st As InkyState = LoadInkyState()
                        ' Recompute upload capability (client may be stale)
                        Dim supportsFilesNow As System.Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            supportsFilesNow = False
                        End Try
                        ' ------------------ (B) File / clipboard object extraction (unchanged logic) ------------
                        Dim extractedDoc As System.String = Nothing
                        Dim extractedLabel As System.String = Nothing
                        Dim attachedType As System.String = Nothing
                        Dim hadInlineExtraction As System.Boolean = False
                        If Not System.String.IsNullOrWhiteSpace(fileObject) Then
                            Dim okOffice As System.Boolean = False
                            Try
                                okOffice = TryExtractOfficeText(fileObject, extractedDoc, extractedLabel)
                            Catch
                                okOffice = False
                            End Try
                            If okOffice Then
                                hadInlineExtraction = True
                                attachedType = "office"
                                fileObject = Nothing    ' Do NOT pass a path to the model
                            Else
                                Dim okText As System.Boolean = False
                                Try
                                    okText = TryExtractTextLike(fileObject, extractedDoc, extractedLabel)
                                Catch
                                    okText = False
                                End Try
                                If okText Then
                                    hadInlineExtraction = True
                                    attachedType = "text"
                                    fileObject = Nothing
                                Else
                                    If (Not supportsFilesNow) AndAlso (Not System.String.IsNullOrWhiteSpace(fileObject)) Then
                                        ' Model does not support file objects
                                        st.History.Add(New ChatTurn With {
                                            .Role = "assistant",
                                            .Markdown = "This model does not support file attachments.",
                                            .Html = MarkdownToHtml("This model does not support file attachments."),
                                            .Utc = System.DateTime.UtcNow
                                        })
                                        SaveInkyState(st)
                                        ' Cleanup temp
                                        Try
                                            If Not System.String.IsNullOrWhiteSpace(uploadedTempPath) AndAlso IO.File.Exists(uploadedTempPath) Then
                                                IO.File.Delete(uploadedTempPath)
                                            End If
                                        Catch
                                        End Try
                                        Return JsonOk(New With {.ok = True, .history = ToBrowserTurns(st.History)})
                                    End If
                                    ' Else: keep fileObject (e.g. PDF) for API that supports raw path/object
                                End If
                            End If
                        End If
                        ' ------------------ (C) Append user turn immediately ------------------
                        Dim userTurn As New ChatTurn With {
                            .Role = "user",
                            .Markdown = textBody,
                            .Html = MarkdownToHtml(textBody),
                            .Utc = Date.UtcNow
                        }
                        st.History.Add(userTurn)
                        ' Cap history (unchanged)
                        Dim cap As Integer = 0
                        Try : cap = INI_ChatCap : Catch : cap = 4000 : End Try
                        Dim clipped As List(Of ChatTurn) = CapHistoryToChars(st, cap)
                        ' Build dialog prompt now (passed into job)
                        Dim sbDialog As New System.Text.StringBuilder()
                        sbDialog.AppendLine("<DIALOG>")
                        For Each t In clipped
                            If t.Role = "user" Then
                                sbDialog.AppendLine("[USER] " & t.Markdown)
                            Else
                                sbDialog.AppendLine("[ASSISTANT] " & t.Markdown)
                            End If
                        Next
                        sbDialog.AppendLine("</DIALOG>")
                        ' Add inline extracted document block if any (unchanged)
                        If hadInlineExtraction AndAlso Not String.IsNullOrWhiteSpace(extractedDoc) Then
                            sbDialog.AppendLine()
                            Dim lbl As String = EscapeForXml(If(extractedLabel, "Attached document"))
                            Dim typ As String = If(String.IsNullOrWhiteSpace(attachedType), "text", attachedType)
                            sbDialog.AppendLine("<ATTACHED_DOCUMENT type=""" & typ & """ label=""" & lbl & """>")
                            sbDialog.AppendLine(extractedDoc)
                            sbDialog.AppendLine("</ATTACHED_DOCUMENT>")
                        End If
                        ' Persist state with user turn BEFORE returning (important)
                        SaveInkyState(st)
                        ' ------------------ (D) Prepare system prompt (same logic) ------------------
                        Dim sysPromptBase As String = _context.SP_Chat.Replace("{Location}", Location)
                        Dim nowLocal As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture)
                        sysPromptBase &= Environment.NewLine & "Current local date/time: " & nowLocal
                        sysPromptBase &= Environment.NewLine & $"Your name is '{AN6}'. "
                        Dim useSecondApiLocal As Boolean = st.UseSecondApi
                        Dim selectedModelKeyLocal As String = st.SelectedModelKey
                        ' Capture file object (may be Nothing after extraction)
                        Dim finalFileObject As String = fileObject
                        Dim tempUploadPathCopy As String = uploadedTempPath
                        ' ------------------ (E) Create background job ------------------
                        Dim jobId As String = Guid.NewGuid().ToString("N")
                        Dim jobCts As New CancellationTokenSource()
                        Dim tcs As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)
                        Dim job As New LlmJob With {
                            .Id = jobId,
                            .CreatedUtc = Date.UtcNow,
                            .Tcs = tcs,
                            .Cts = jobCts,
                            .UseSecond = useSecondApiLocal,
                            .FileObject = finalFileObject
                        }
                        If Not jobMap.TryAdd(jobId, job) Then
                            jobCts.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)
                        ' Keep old global CTS reference for legacy inky_cancel support
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCts
                        ' Replace the original System.Threading.Tasks.Task.Run(Sub() ... End Sub) block with this:
                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim configSnapshot As ModelConfig = Nothing
                                Dim snapshotTaken As Boolean = False
                                Dim alternateApplied As Boolean = False
                                Dim localOutput As String = ""
                                Try
                                    ' (1) Alternate model application (safer pattern)
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            SyncLock AlternateModelLock
                                                configSnapshot = GetCurrentConfig(_context)
                                                snapshotTaken = True
                                                Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                                Dim sel = alts?.FirstOrDefault(
                                                    Function(m)
                                                        If m Is Nothing Then Return False
                                                        If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                           String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                        If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                           String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                        Return False
                                                    End Function)
                                                If sel IsNot Nothing Then
                                                    ApplyModelConfig(_context, sel)
                                                    alternateApplied = True
                                                End If
                                            End SyncLock
                                        Catch
                                            ' Swallow – we will still attempt restore if snapshotTaken
                                        End Try
                                    End If
                                    ' (2) Run LLM
                                    localOutput = RunLlmAsync(
                                        sysPromptBase,
                                        sbDialog.ToString(),
                                        useSecondApiLocal,
                                        False,
                                        finalFileObject,
                                        jobCts.Token
                                    ).GetAwaiter().GetResult()
                                    If localOutput Is Nothing Then localOutput = String.Empty
                                    localOutput = SanitizeModelOutputForBrowser(localOutput).Trim()
                                    If localOutput.Length > 0 AndAlso
                                       localOutput.Equals("Operation was canceled by the user.", StringComparison.OrdinalIgnoreCase) Then
                                        localOutput = "Aborted by user."
                                    End If
                                    ' (3) Build assistant turn or error turn
                                    Dim assistantText As String = localOutput
                                    Dim wasCanceled As Boolean = jobCts.IsCancellationRequested
                                    If assistantText.Length = 0 Then
                                        assistantText = If(wasCanceled,
                                                           "Aborted by user.",
                                                           "Error: The model did not provide a response.")
                                    End If
                                    Dim htmlOut = MarkdownToHtml(assistantText)
                                    ' Reload latest state (in case user switched chats meanwhile)
                                    Dim stJob = LoadInkyState()
                                    stJob.History.Add(New ChatTurn With {
                                        .Role = "assistant",
                                        .Markdown = assistantText,
                                        .Html = htmlOut,
                                        .Utc = Date.UtcNow
                                    })
                                    stJob.LastAssistantText = assistantText
                                    SaveInkyState(stJob)
                                    If wasCanceled AndAlso localOutput.Length = 0 Then
                                        tcs.TrySetCanceled()
                                    Else
                                        tcs.TrySetResult(assistantText)
                                    End If
                                Catch exOp As OperationCanceledException
                                    tcs.TrySetCanceled()
                                Catch ex As System.Exception
                                    tcs.TrySetException(ex)
                                Finally
                                    ' (4) Cleanup temp upload
                                    Try
                                        If Not String.IsNullOrWhiteSpace(tempUploadPathCopy) AndAlso IO.File.Exists(tempUploadPathCopy) Then
                                            IO.File.Delete(tempUploadPathCopy)
                                        End If
                                    Catch
                                    End Try
                                    ' (5) Restore original config unconditionally if snapshot was taken
                                    Try
                                        If snapshotTaken Then
                                            SyncLock AlternateModelLock
                                                RestoreDefaults(_context, configSnapshot)
                                            End SyncLock
                                        End If
                                    Catch
                                        ' Ignore restore errors
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)
                        ' ------------------ (F) Immediate response with job id ------------------
                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobId,
                            .status = "running",
                            .history = ToBrowserTurns(st.History)
                        })

                    Case "inky_pure"
                        ' (Pure mode logic preserved)
                        ' ------------------------------------------------------------------
                        Dim fileObject As String = j("FileObject")?.ToString()
                        Dim textBody As String = j("Text")?.ToString()
                        If String.IsNullOrWhiteSpace(textBody) Then
                            Return JsonErr("Please enter a message.")
                        End If
                        Dim stPure As InkyState = LoadInkyState()
                        ' Decide model context (still respect selected model choice)
                        Dim useSecondApiLocal As Boolean = stPure.UseSecondApi
                        Dim selectedModelKeyLocal As String = stPure.SelectedModelKey
                        ' Validate file object support (only pass through if supported; no extraction)
                        Dim supportsFilesNow As Boolean = False
                        Try
                            supportsFilesNow = ComputeSupportsFiles(useSecondApiLocal, selectedModelKeyLocal)
                        Catch
                            supportsFilesNow = False
                        End Try
                        If Not supportsFilesNow Then
                            fileObject = Nothing
                        End If
                        ' Record user turn (prefixed)
                        stPure.History.Add(New ChatTurn With {
                            .Role = "user",
                            .Markdown = "Pure: " & textBody,
                            .Html = MarkdownToHtml("Pure: " & textBody),
                            .Utc = Date.UtcNow
                        })
                        SaveInkyState(stPure)
                        ' Prepare background job (mirrors inky_send logic but no dialog/system prompt)
                        Dim jobIdP As String = Guid.NewGuid().ToString("N")
                        Dim jobCtsP As New CancellationTokenSource()
                        Dim tcsP As New TaskCompletionSource(Of String)(TaskCreationOptions.RunContinuationsAsynchronously)
                        Dim jobP As New LlmJob With {
                            .Id = jobIdP,
                            .CreatedUtc = Date.UtcNow,
                            .Tcs = tcsP,
                            .Cts = jobCtsP,
                            .UseSecond = useSecondApiLocal,
                            .FileObject = fileObject
                        }
                        If Not jobMap.TryAdd(jobIdP, jobP) Then
                            jobCtsP.Dispose()
                            Return JsonErr("Failed to register job.")
                        End If
                        Threading.Interlocked.Increment(activeJobs)
                        ' Maintain legacy cancellation
                        Try
                            If llmOperationCts IsNot Nothing Then
                                llmOperationCts.Cancel()
                                llmOperationCts.Dispose()
                            End If
                        Catch
                        End Try
                        llmOperationCts = jobCtsP
                        System.Threading.Tasks.Task.Run(
                            Sub()
                                Dim originalCfgLoadedP As Boolean = False
                                Dim usedAlternate As Boolean = False
                                Try
                                    ' Apply alternate model if a specific alternate is selected
                                    If useSecondApiLocal AndAlso Not String.IsNullOrWhiteSpace(selectedModelKeyLocal) Then
                                        Try
                                            originalConfig = GetCurrentConfig(_context) : originalCfgLoadedP = True
                                            Dim alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                                            Dim sel = alts?.FirstOrDefault(
                                                Function(m)
                                                    If m Is Nothing Then Return False
                                                    If Not String.IsNullOrWhiteSpace(m.ModelDescription) AndAlso
                                                       String.Equals(m.ModelDescription, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    If Not String.IsNullOrWhiteSpace(m.Model) AndAlso
                                                       String.Equals(m.Model, selectedModelKeyLocal, StringComparison.OrdinalIgnoreCase) Then Return True
                                                    Return False
                                                End Function)
                                            If sel IsNot Nothing Then
                                                ApplyModelConfig(_context, sel)
                                                usedAlternate = True
                                            End If
                                        Catch
                                        End Try
                                    End If
                                    ' Raw call: NO system prompt, NO history packaging
                                    Dim output = RunLlmAsync("", textBody, useSecondApiLocal, False, fileObject, jobCtsP.Token).GetAwaiter().GetResult()
                                    If output Is Nothing Then output = String.Empty
                                    output = SanitizeModelOutputForBrowser(output).Trim()
                                    If jobCtsP.IsCancellationRequested AndAlso output.Length = 0 Then
                                        tcsP.TrySetCanceled()
                                    Else
                                        If output.Length = 0 Then
                                            output = If(jobCtsP.IsCancellationRequested, "Aborted by user.", "Error: The model returned no content.")
                                        End If
                                        Dim stFin = LoadInkyState()
                                        stFin.History.Add(New ChatTurn With {
                                            .Role = "assistant",
                                            .Markdown = output,
                                            .Html = MarkdownToHtml(output),
                                            .Utc = Date.UtcNow
                                        })
                                        stFin.LastAssistantText = output
                                        SaveInkyState(stFin)
                                        tcsP.TrySetResult(output)
                                    End If
                                Catch exOp As OperationCanceledException
                                    tcsP.TrySetCanceled()
                                Catch ex As System.Exception
                                    tcsP.TrySetException(ex)
                                Finally
                                    ' Restore config if alternate used
                                    Try
                                        If useSecondApiLocal AndAlso (usedAlternate OrElse originalCfgLoadedP) AndAlso originalConfigLoaded Then
                                            RestoreDefaults(_context, originalConfig)
                                            originalConfigLoaded = False
                                        End If
                                    Catch
                                    End Try
                                    Threading.Interlocked.Decrement(activeJobs)
                                End Try
                            End Sub)
                        Return JsonOk(New With {
                            .ok = True,
                            .job = jobIdP,
                            .status = "running",
                            .history = ToBrowserTurns(stPure.History)
                        })


                    Case "inky_jobstatus"
                        ' (Job status check preserved)
                        ' ------------------------------------------------------------------
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then
                            Return JsonErr("Unknown job.")
                        End If
                        ' Cleanup old jobs (lazy TTL)
                        If (Date.UtcNow - job.CreatedUtc).TotalMinutes > JobTtlMinutes Then
                            Dim dump As LlmJob = Nothing
                            jobMap.TryRemove(jobId, dump)
                            Return JsonErr("Job expired.")
                        End If
                        Dim t = job.Tcs.Task
                        If Not t.IsCompleted Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "running"})
                        End If
                        If t.IsCanceled Then
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "canceled"})
                        ElseIf t.IsFaulted Then
                            Return JsonOk(New With {.ok = False, .job = jobId, .status = "error", .error = t.Exception.GetBaseException().Message})
                        Else
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "done", .result = t.Result})
                        End If

                    Case "inky_canceljob"
                        ' (Direct cancel by id)
                        ' ------------------------------------------------------------------
                        Dim jobId = j("Job")?.ToString()
                        If String.IsNullOrWhiteSpace(jobId) Then Return JsonErr("Missing Job id.")
                        Dim job As LlmJob = Nothing
                        If Not jobMap.TryGetValue(jobId, job) Then Return JsonErr("Unknown job.")
                        Try
                            job.Cts.Cancel()
                            Return JsonOk(New With {.ok = True, .job = jobId, .status = "cancelRequested"})
                        Catch ex As System.Exception
                            Return JsonErr("Cancel failed: " & ex.Message)
                        End Try

                    Case "inky_toword"
                        Try
                            Dim st As InkyState = LoadInkyState()
                            Dim turns = If(st?.History, New System.Collections.Generic.List(Of ChatTurn)())

                            Dim sb As New System.Text.StringBuilder()
                            sb.AppendLine("# " & GetBotName() & " — Chat thread")
                            sb.AppendLine()
                            sb.AppendLine("Generated: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture))
                            sb.AppendLine()

                            For Each t In turns
                                If t Is Nothing Then Continue For
                                Dim who As String = If(String.Equals(t.Role, "user", StringComparison.OrdinalIgnoreCase), "You", GetBotName())
                                sb.AppendLine("## " & who)
                                sb.AppendLine(t.Markdown)
                                sb.AppendLine()
                            Next

                            Dim markdownThread As String = sb.ToString()

                            Dim ok As Boolean = Await TryCreateWordDocFromMarkdown(markdownThread).ConfigureAwait(False)

                            If Not ok Then Return JsonErr("Could not create a Word document.")

                            Return JsonOk(New With {.ok = True})
                        Catch ex As System.Exception
                            Return JsonErr("To Word failed: " & ex.Message)
                        End Try

                    Case "inky_clear"
                        ' (Clear history)
                        ' ------------------------------------------------------------------
                        Dim stClear As New InkyState()
                        SaveInkyState(stClear)
                        Return JsonOk(New With {
                            .ok = True,
                            .activeChat = activeChatId,
                            .greeting = GetFriendlyGreeting()
                        })

                    Case "inky_copylast"
                        ' (Copy last assistant response)
                        ' ------------------------------------------------------------------
                        Dim stCopy As InkyState = LoadInkyState()
                        If System.String.IsNullOrWhiteSpace(stCopy.LastAssistantText) Then
                            Return JsonErr("No assistant response available to copy.")
                        End If
                        Await SwitchToUi(Sub()
                                             SLib.PutInClipboard(MarkdownToRtfConverter.Convert(stCopy.LastAssistantText))
                                         End Sub).ConfigureAwait(False)
                        Return JsonOk(New With {.ok = True})

                    Case "inky_setmodel"
                        ' (Model selection)
                        ' ------------------------------------------------------------------
                        Dim key As System.String = j("Key")?.ToString()
                        Dim st As InkyState = LoadInkyState()
                        If System.String.IsNullOrWhiteSpace(key) OrElse System.String.Equals(key, "default", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = False
                            st.SelectedModelKey = ""
                        ElseIf System.String.Equals(key, "__second__", System.StringComparison.OrdinalIgnoreCase) Then
                            st.UseSecondApi = True
                            st.SelectedModelKey = ""
                        Else
                            st.UseSecondApi = True
                            st.SelectedModelKey = key
                        End If
                        Try
                            My.Settings.Inky_UseSecondApiSelected = st.UseSecondApi
                            My.Settings.Inky_SelectedModelKey = st.SelectedModelKey
                            My.Settings.Save()
                        Catch
                        End Try
                        ' Re-evaluate upload capability for selected model
                        Try
                            st.SupportsFileUploads = ComputeSupportsFiles(st.UseSecondApi, st.SelectedModelKey)
                        Catch
                            st.SupportsFileUploads = False
                        End Try
                        SaveInkyState(st)
                        Return JsonOk(New With {.ok = True, .supportsFiles = st.SupportsFileUploads, .activeChat = activeChatId})

                    Case "inky_toggletheme"
                        ' (Theme toggle)
                        ' ------------------------------------------------------------------
                        Dim st As InkyState = LoadInkyState()
                        st.DarkMode = Not st.DarkMode
                        SaveInkyState(st)
                        Try
                            My.Settings.Inky_DarkMode = st.DarkMode
                            My.Settings.Save()
                        Catch
                        End Try
                        Return JsonOk(New With {.ok = True, .darkMode = st.DarkMode, .activeChat = activeChatId})

                    Case Else
                        Return JsonErr("Unknown command.")
                End Select

            Catch ex As System.Exception
                Return JsonErr("Bad request: " & ex.Message)
            End Try
        End If

        ' ---- FALLBACK to existing command dispatcher (legacy) ----
        Dim j0 = Newtonsoft.Json.Linq.JObject.Parse(If(body, "{}"))
        Dim cmd0 = j0("Command")?.ToString()
        Dim textBody0 = j0("Text")?.ToString()
        Dim sourceUrl = j0("URL")?.ToString()

        Select Case cmd0
            Case "redink_sendtooutlook"
                ' (Legacy Outlook insert)
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                ' All Outlook automation on UI thread
                Await SwitchToUi(Sub()
                                     Dim olApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
                                     Dim insp As Microsoft.Office.Interop.Outlook.Inspector = ComRetry(Function() olApp.ActiveInspector())
                                     If insp Is Nothing Then Exit Sub
                                     ' Guard CurrentItem (never access inline)
                                     Dim curr As Object = Nothing
                                     Try
                                         curr = ComRetry(Function() insp.CurrentItem)
                                     Catch
                                         curr = Nothing
                                     End Try
                                     If curr Is Nothing OrElse Not TypeOf curr Is Microsoft.Office.Interop.Outlook.MailItem Then
                                         Exit Sub
                                     End If
                                     Dim mail As Microsoft.Office.Interop.Outlook.MailItem =
                                         CType(curr, Microsoft.Office.Interop.Outlook.MailItem)
                                     ' Guard Sent property
                                     Try
                                         If ComRetry(Function() mail.Sent) Then
                                             If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                             Exit Sub
                                         End If
                                     Catch
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                         Exit Sub
                                     End Try
                                     ' Guard WordEditor and selection
                                     Dim doc As Microsoft.Office.Interop.Word.Document = Nothing
                                     Try
                                         doc = ComRetry(Function() CType(insp.WordEditor, Microsoft.Office.Interop.Word.Document))
                                         If doc Is Nothing Then Exit Sub
                                         Dim rng As Microsoft.Office.Interop.Word.Range = Nothing
                                         Try
                                             rng = doc.Application.Selection.Range
                                             doc.Application.ScreenUpdating = False
                                             rng.Text = textBody0 & " (" & sourceUrl & ")"
                                             doc.Application.ScreenUpdating = True
                                         Finally
                                             If rng IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(rng) : rng = Nothing
                                         End Try
                                     Finally
                                         If doc IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(doc) : doc = Nothing
                                         If insp IsNot Nothing Then System.Runtime.InteropServices.Marshal.ReleaseComObject(insp) : insp = Nothing
                                     End Try
                                 End Sub)
                Return ""

            Case "redink_translate"
                ' ─── 1  guard clauses ─────────────────────────────────────────
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                Dim targetLang As String = Await SwitchToUi(Function()
                                                                Return SLib.ShowCustomInputBox(
                                                                    "Enter your target language:",
                                                                    AN & " Translate (for Browser)",
                                                                    True, INI_Language1)
                                                            End Function)
                If String.IsNullOrWhiteSpace(targetLang) OrElse targetLang = "ESC" Then
                    Return ""
                End If
                TranslateLanguage = targetLang.Trim()
                ' ─── 2  call the LLM on the UI thread, get Task(Of String) ─────
                Dim llmOut As String = Await RunLlmAsync(
                    InterpolateAtRuntime(SP_Translate),
                    $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")
                ' ─── 3  clean up the wrapper tags / markdown ──────────────────
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "") _
                       .Replace("</TEXTTOPROCESS>", "") _
                       .Replace("**", "").Trim()
                If llmOut = "" Then Return ""
                ' Optional: copy to clipboard so the user can paste manually
                Await SwitchToUi(Sub() SLib.PutInClipboard(llmOut))
                ' ─── 4  SEND the translation back to the caller ───────────────
                Return llmOut

            Case "redink_correct"
                If String.IsNullOrWhiteSpace(textBody0) Then Return ""
                ' 1)  Run the LLM on the UI thread
                Dim llmOut As String = Await RunLlmAsync(
                    InterpolateAtRuntime(SP_Correct),
                    $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>")
                llmOut = llmOut.Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "")
                If llmOut = "" Then Return ""
                ' 2)  Show the compare / preview window (synchronous)
                Await SwitchToUi(Sub()
                                     CompareAndInsertText(textBody0, llmOut, True)
                                 End Sub)
                ' 3)  
                Dim accepted As Boolean = Await WaitForPreviewDecisionAsync()
                If Not accepted Then Return ""
                Return llmOut

            Case "redink_freestyle"
                ' (Freestyle flow preserved)
                '─── A  gather prompt on UI thread ──────────────────────────────
                ' ---------------------------------------------------------------
                Dim noText As Boolean = String.IsNullOrWhiteSpace(textBody0)
                Dim promptCaption As String = AN & " Freestyle (for Browser)"
                Dim wordInstalled As Boolean = False
                Try
                    Dim wordApp As Object = CreateObject("Word.Application")
                    wordInstalled = True
                    wordApp.Quit()
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp)
                Catch ex As System.Exception
                    wordInstalled = False
                End Try
                Dim sb As New System.Text.StringBuilder()
                If noText Then
                    sb.Append("Please provide the prompt you wish to execute ")
                Else
                    sb.Append("Please provide the prompt you wish to execute using the selected text ")
                End If
                sb.Append("('" & MarkupPrefix & "' for markups, '" & InsertPrefix & "' for direct insert" & If(wordInstalled, " and '" & NewDocPrefix & "' to put the output in a new Word document)", ")"))
                If INI_PromptLib Then sb.Append(" or press 'OK' for the prompt library")
                If INI_SecondAPI Then sb.Append($"; add '{SecondAPICode}' to use {If(String.IsNullOrWhiteSpace(INI_AlternateModelPath), $"the secondary model ({INI_Model_2})", "one of the other models")}")
                If Not String.IsNullOrWhiteSpace(My.Settings.LastPrompt) Then sb.Append("; ctrl-p for your last prompt")
                sb.Append(":")
                Dim promptMsg As String = sb.ToString()
                Dim OptionalButtons As System.Tuple(Of String, String, String)()
                If wordInstalled Then
                    OptionalButtons = {
                        System.Tuple.Create("OK, do a new doc", $"Use this to automatically insert '{NewDocPrefix}' as a prefix.", NewDocPrefix)
                    }
                End If
                OtherPrompt = Await SwitchToUi(Function()
                                                   Return SLib.ShowCustomInputBox(promptMsg, promptCaption, False, "", My.Settings.LastPrompt, If(wordInstalled, OptionalButtons, Nothing))
                                               End Function)
                Dim doMarkupFlag As Boolean = False
                Dim doInsertFlag As Boolean = False
                Dim UseSecondAPI As Boolean = False
                Dim DoNewDoc As Boolean = False
                If String.IsNullOrEmpty(OtherPrompt) AndAlso OtherPrompt <> "ESC" AndAlso INI_PromptLib Then
                    Dim sel = Await SwitchToUi(Function()
                                                   Return ShowPromptSelector(INI_PromptLibPath, INI_PromptLibPathLocal, Not noText, Nothing)
                                               End Function)
                    OtherPrompt = sel.Item1
                    doMarkupFlag = sel.Item2
                    doInsertFlag = Not sel.Item4
                End If
                If String.IsNullOrWhiteSpace(OtherPrompt) OrElse OtherPrompt = "ESC" Then
                    Return ""
                End If
                My.Settings.LastPrompt = OtherPrompt
                My.Settings.Save()
                If OtherPrompt.StartsWith(InsertPrefix, StringComparison.OrdinalIgnoreCase) Then
                    OtherPrompt = OtherPrompt.Substring(InsertPrefix.Length).Trim()
                    doInsertFlag = True
                ElseIf OtherPrompt.StartsWith(MarkupPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(MarkupPrefix.Length).Trim()
                    doMarkupFlag = True
                    doInsertFlag = True
                ElseIf OtherPrompt.StartsWith(NewDocPrefix, StringComparison.OrdinalIgnoreCase) AndAlso Not noText Then
                    OtherPrompt = OtherPrompt.Substring(NewDocPrefix.Length).Trim()
                    DoNewDoc = True
                    doMarkupFlag = False
                End If
                If INI_SecondAPI AndAlso OtherPrompt.Contains(SecondAPICode) Then
                    UseSecondAPI = True
                    OtherPrompt = OtherPrompt.Replace(SecondAPICode, "").Trim()
                    If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                        Dim sel = Await SwitchToUi(Function()
                                                       Return Not ShowModelSelection(_context, INI_AlternateModelPath)
                                                   End Function)
                        If sel Then
                            originalConfigLoaded = False
                            Return ""
                        End If
                    End If
                End If
                '─── B  call the LLM on UI thread (async) ──────────────────────
                Dim llmResult As String
                If noText Then
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleNoText), "", UseSecondAPI)
                Else
                    llmResult = Await RunLlmAsync(InterpolateAtRuntime(SP_FreestyleText), $"<TEXTTOPROCESS>{textBody0}</TEXTTOPROCESS>", UseSecondAPI)
                End If
                llmResult = llmResult.Replace("<TEXTTOPROCESS>", "") _
                             .Replace("</TEXTTOPROCESS>", "") _
                             .Trim()
                If String.IsNullOrEmpty(llmResult) Then Return ""
                '─── C  present / insert / clipboard exactly like old code ─────
                ' A) markup path (implies insert)  -----------------------------
                If doMarkupFlag Then
                    Await SwitchToUi(Sub()
                                         CompareAndInsertText(textBody0, llmResult, True)
                                     End Sub)
                    Dim acceptedFs As Boolean = Await WaitForPreviewDecisionAsync()
                    If Not acceptedFs Then Return ""
                    Return llmResult
                End If
                ' B) plain insert path  ----------------------------------------
                If doInsertFlag Then
                    Return llmResult
                End If
                If DoNewDoc And wordInstalled Then
                    If Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    End If
                    Await SwitchToUi(Sub()
                                         ShowCustomMessageBox("Could not create new Word document and insert the LLM output; providing your output to a separate window.")
                                     End Sub)
                End If
                ' C) clipboard-only path  --------------------------------------
                Dim finalTxt As String = Await SwitchToUi(Function()
                                                              Return SLib.ShowCustomWindow(
                                                                  "The LLM has provided the following result (you can edit it):",
                                                                  llmResult,
                                                                  "You can choose whether you want to have the original text put into the clipboard or your text with any changes you have made (without formatting)." & If(wordInstalled, " If you choose to insert the original text with formatting, a new Word document will be created with it. ", " ") & "If you select Cancel, nothing will be put into the clipboard (you can yourself copy it to the clipboard).",
                                                                  AN, False, True, If(wordInstalled, True, False))
                                                          End Function)
                If String.Equals(finalTxt, "Markdown", StringComparison.OrdinalIgnoreCase) Then
                    If wordInstalled AndAlso Await TryCreateWordDocFromMarkdown(llmResult) Then
                        Return ""
                    Else
                        Await SwitchToUi(Sub()
                                             ShowCustomMessageBox("Could not create new Word document and insert the LLM output (however, it will be in the clipboard).")
                                             finalTxt = MarkdownToRtfConverter.Convert(llmResult)
                                         End Sub)
                    End If
                End If
                If Not String.IsNullOrWhiteSpace(finalTxt) Then
                    Await SwitchToUi(Sub() SLib.PutInClipboard(finalTxt))
                End If

                Dim plainText As String

                ' Convert RTF to plain text using RichTextBox
                Using rtb As New RichTextBox()
                    rtb.Rtf = finalTxt
                    plainText = rtb.Text
                End Using

                Return plainText
        End Select

        Return ""
    End Function

    ''' <summary>
    ''' Creates a new Word document, inserts markdown-converted content, shows user message. Returns True on success.
    ''' </summary>
    Private Async Function TryCreateWordDocFromMarkdown(markdown As String) As Task(Of Boolean)
        Try
            Dim wordApp As New Microsoft.Office.Interop.Word.Application()
            wordApp.Visible = True
            Dim newDoc As Microsoft.Office.Interop.Word.Document = wordApp.Documents.Add()
            Dim docSelection As Microsoft.Office.Interop.Word.Selection = wordApp.Selection
            InsertTextWithMarkdown(docSelection, markdown, True)
            Await SwitchToUi(Sub()
                                 ShowCustomMessageBox("Your Word document has been created. It may be hidden behind the other windows.")
                             End Sub)
            Return True
        Catch
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Builds selectable model list (primary, secondary, alternatives) based on current configuration and persisted state.
    ''' </summary>
    Private Async Function GetModelListForBrowserAsync(ByVal st As InkyState) _
        As System.Threading.Tasks.Task(Of System.Collections.Generic.List(Of Object))

        Dim list As New System.Collections.Generic.List(Of Object)()

        ' 0) If config is not loaded or primary name is blank, try to (re)load once
        If (Not INIloaded) OrElse (String.IsNullOrWhiteSpace(INI_Model) AndAlso Not INI_SecondAPI) Then
            Try
                InitializeConfig(False, False)
            Catch
            End Try
        End If

        ' 1) Reconcile persisted preference
        Try
            Dim savedSecond As Boolean = My.Settings.Inky_UseSecondApiSelected
            Dim savedKey As String = My.Settings.Inky_SelectedModelKey
            Dim shouldApply As Boolean =
                (savedSecond <> st.UseSecondApi) OrElse
                (savedSecond AndAlso Not String.Equals(savedKey, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase))
            If shouldApply Then
                st.UseSecondApi = savedSecond
                st.SelectedModelKey = If(savedSecond, savedKey, "")
                SaveInkyState(st)
            End If
        Catch
        End Try

        ' 2) Availability
        Dim hasPrimary As Boolean = Not String.IsNullOrWhiteSpace(INI_Model)
        Dim hasSecondApi As Boolean = INI_SecondAPI
        Dim hasSecondModelName As Boolean = Not String.IsNullOrWhiteSpace(INI_Model_2)
        Dim hasSecondary As Boolean = hasSecondApi AndAlso hasSecondModelName

        ' Alt list
        Dim alts As System.Collections.Generic.List(Of SharedLibrary.SharedLibrary.ModelConfig) = Nothing
        Dim altCount As Integer = 0
        Try
            If hasSecondApi AndAlso Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                alts = LoadAlternativeModels(INI_AlternateModelPath, _context)
                If alts IsNot Nothing Then altCount = alts.Count
            End If
        Catch
            altCount = 0
            alts = Nothing
        End Try

        ' 3) Normalize stale alternate selection EVEN if alts could not be loaded
        If st.UseSecondApi AndAlso Not String.IsNullOrWhiteSpace(st.SelectedModelKey) Then
            Dim exists As Boolean = False
            If alts IsNot Nothing Then
                exists = alts.Any(Function(m)
                                      If m Is Nothing Then Return False
                                      Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                                      Return String.Equals(label, st.SelectedModelKey, StringComparison.OrdinalIgnoreCase)
                                  End Function)
            Else
                exists = False
            End If
            If Not exists Then
                st.SelectedModelKey = ""
                SaveInkyState(st)
            End If
        End If

        ' 4) Only primary
        If hasPrimary AndAlso Not hasSecondary AndAlso altCount = 0 Then
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
            Return list
        End If

        ' 5) Build list
        If hasPrimary Then
            list.Add(New With {.key = "__hdr_primary__", .label = "Primary model:", .selected = False, .disabled = True, .isSeparator = False})
            list.Add(New With {
                .key = "default",
                .label = INI_Model,
                .selected = (Not st.UseSecondApi),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If hasSecondary Then
            list.Add(New With {.key = "__hdr_secondary__", .label = "Secondary model:", .selected = False, .disabled = True, .isSeparator = False})
            list.Add(New With {
                .key = "__second__",
                .label = INI_Model_2,
                .selected = (st.UseSecondApi AndAlso String.IsNullOrWhiteSpace(st.SelectedModelKey)),
                .disabled = False,
                .isSeparator = False
            })
        End If

        If altCount > 0 AndAlso alts IsNot Nothing Then
            list.Add(New With {.key = "__sep__", .label = "Alternative models:", .selected = False, .disabled = True, .isSeparator = True})
            For Each m In alts
                If m Is Nothing Then Continue For
                Dim label = If(Not String.IsNullOrWhiteSpace(m.ModelDescription), m.ModelDescription, m.Model)
                If String.IsNullOrWhiteSpace(label) Then label = "Model"
                list.Add(New With {
                    .key = label,
                    .label = label,
                    .selected = (st.UseSecondApi AndAlso String.Equals(st.SelectedModelKey, label, StringComparison.OrdinalIgnoreCase)),
                    .disabled = False,
                    .isSeparator = False
                })
            Next
        End If

        ' 6) Never return an empty list → synthesize a safe primary entry
        If list.Count = 0 Then
            Dim lbl = If(String.IsNullOrWhiteSpace(INI_Model), "Primary model (not configured)", INI_Model)
            list.Add(New With {
                .key = "default",
                .label = lbl,
                .selected = True,
                .disabled = String.IsNullOrWhiteSpace(INI_Model),
                .isSeparator = False
            })
        End If

        Return list
    End Function


    ''' <summary>
    ''' Sanitizes raw model output by removing role markers and normalizing excessive blank lines.
    ''' </summary>
    Private Function SanitizeModelOutputForBrowser(ByVal raw As System.String) As System.String
        If raw Is Nothing Then Return System.String.Empty

        Dim s As System.String = raw

        ' 1) Remove full role-only lines (e.g. "[ASSISTANT]" or "[USER]:")
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*:?\s*$",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 2) Remove role markers at line start (preserve text)
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*\[(?:assistant|user)\]\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 3) Remove alternative "ASSISTANT:" / "USER:" prefixes at line start
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(?im)^\s*(?:assistant|user)\s*:\s*",
            "",
            System.Text.RegularExpressions.RegexOptions.None)

        ' 4) Normalize excessive blank lines
        s = System.Text.RegularExpressions.Regex.Replace(
            s,
            "(\r?\n){3,}",
            System.Environment.NewLine & System.Environment.NewLine,
            System.Text.RegularExpressions.RegexOptions.None)

        Return s
    End Function

    ''' <summary>
    ''' Maps internal ChatTurn list to browser DTO objects (camelCase).
    ''' </summary>
    Private Function ToBrowserTurns(list As System.Collections.Generic.List(Of ChatTurn)) _
        As System.Collections.Generic.List(Of Object)

        Dim out As New System.Collections.Generic.List(Of Object)()
        For Each t In list
            out.Add(New With {
                .role = t.Role,
                .markdown = t.Markdown,
                .html = t.Html,
                .utc = t.Utc
            })
        Next
        Return out
    End Function

    ''' <summary>
    ''' Invokes user32.dll to obtain the current process GDI (uiFlags=0) or USER (uiFlags=1) handle count.
    ''' </summary>
    ''' <param name="hProcess">Handle of the process being queried.</param>
    ''' <param name="uiFlags">0 for GDI handles, 1 for USER handles.</param>
    ''' <returns>The handle count reported by the OS for the specified resource type.</returns>

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetGuiResources(hProcess As System.IntPtr, uiFlags As System.Int32) As System.UInt32
    End Function

    ''' <summary>
    ''' Returns current process GDI handle count.
    ''' </summary>
    Private Shared Function GetGdiCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 0UI)
    End Function

    ''' <summary>
    ''' Returns current process USER handle count.
    ''' </summary>
    Private Shared Function GetUserCount() As System.UInt32
        Return GetGuiResources(System.Diagnostics.Process.GetCurrentProcess().Handle, 1UI)
    End Function

End Class