' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland.
' All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: DrawioEditorForm.vb
' Summary:
'   Hosts the diagrams.net (draw.io) embedded editor inside a WinForms window
'   using WebView2. Loads a diagram from an mxfile XML string and saves edited
'   XML back to a local .drawio file. Supports export via the diagrams.net embed
'   postMessage API.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Web.WebView2.Core
Imports Microsoft.Web.WebView2.WinForms
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Class DrawioEditorForm
    Inherits Form

    Private WithEvents webView As WebView2

    Private ReadOnly _xmlContent As String
    Private ReadOnly _saveFilePath As String
    Private ReadOnly _fileName As String

    Private Const DRAWIO_EMBED_URL As String =
        "https://embed.diagrams.net/?embed=1&proto=json&spin=1&saveAndExit=1&noSaveBtn=0"

    Private ReadOnly _disableInternetAfterLoad As Boolean
    Private _offlineModeEnabled As Boolean
    Private _networkFilterInstalled As Boolean
    Private _navigationCompletedOnce As Boolean

    Private _pendingExportFormat As String
    Private _queuedExportFormat As String
    Private _hostReadyRaised As Boolean
    Private _hostIsReady As Boolean

    ''' <summary>
    ''' Raised when the HTML host page has initialized and can message the WebView host.
    ''' </summary>
    Public Event HostReady As EventHandler

    ''' <summary>
    ''' Creates a new editor window for draw.io content loaded from an XML string.
    ''' </summary>
    ''' <param name="xmlContent">The mxfile XML to load into the editor.</param>
    ''' <param name="saveFilePath">The local path to write the updated diagram XML to.</param>
    ''' <param name="disableInternetAfterLoad">
    ''' If <see langword="True"/>, http/https requests from within this WebView2 instance are blocked after draw.io has accepted the initial load.
    ''' </param>
    Public Sub New(xmlContent As String, saveFilePath As String, Optional disableInternetAfterLoad As Boolean = False)
        _xmlContent = If(xmlContent, "")
        _saveFilePath = saveFilePath
        _fileName = Path.GetFileNameWithoutExtension(saveFilePath)
        _disableInternetAfterLoad = disableInternetAfterLoad

        InitializeComponent()
        InitializeWebViewAsync()
    End Sub

    ''' <summary>
    ''' Initializes the form UI and hosting WebView2 control.
    ''' </summary>
    Private Sub InitializeComponent()
        Me.Text = $"{AN} - Draw.io Diagram Editor"
        Me.Width = 1400
        Me.Height = 900
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.Sizable

        Try
            Dim bmp As New Bitmap(GetLogoBitmap(LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        webView = New WebView2() With {.Dock = DockStyle.Fill}
        Me.Controls.Add(webView)
    End Sub

    ''' <summary>
    ''' Creates the WebView2 environment, wires events, installs network gating, and navigates to the local host HTML.
    ''' </summary>
    Private Async Sub InitializeWebViewAsync()
        Try
            Dim userDataFolder As String = Path.Combine(Path.GetTempPath(), "RedInk_WebView2_Drawio")
            Directory.CreateDirectory(userDataFolder)

            Dim env As CoreWebView2Environment =
                Await CoreWebView2Environment.CreateAsync(Nothing, userDataFolder)

            Await webView.EnsureCoreWebView2Async(env)

            webView.CoreWebView2.Settings.AreDevToolsEnabled = True
            webView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = True

            AddHandler webView.CoreWebView2.WebMessageReceived, AddressOf OnWebMessageReceived
            AddHandler webView.CoreWebView2.NavigationCompleted, AddressOf OnNavigationCompleted

            InstallOfflineNetworkFilter()

            Dim html As String = BuildHostHtml(_xmlContent, _fileName)

            Dim hostDir As String = Path.Combine(Path.GetTempPath(), "RedInk_DrawioHost")
            Directory.CreateDirectory(hostDir)

            Dim hostFilePath As String = Path.Combine(hostDir, "drawio_host.html")
            File.WriteAllText(hostFilePath, html, System.Text.Encoding.UTF8)

            Dim hostUri As New Uri(hostFilePath)
            webView.CoreWebView2.Navigate(hostUri.AbsoluteUri)

        Catch ex As Exception
            MessageBox.Show(
                $"Could not initialize the diagram editor: {ex.Message}{vbCrLf}{vbCrLf}" &
                $"Your diagram has been saved to:{vbCrLf}{_saveFilePath}",
                $"{AN} - Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
            Me.Close()
        End Try
    End Sub

    ''' <summary>
    ''' Installs a WebView2 request filter used to block http/https requests when offline mode is enabled.
    ''' </summary>
    Private Sub InstallOfflineNetworkFilter()
        If webView.CoreWebView2 Is Nothing Then Return
        If _networkFilterInstalled Then Return

        _networkFilterInstalled = True
        webView.CoreWebView2.AddWebResourceRequestedFilter("*", CoreWebView2WebResourceContext.All)
        AddHandler webView.CoreWebView2.WebResourceRequested, AddressOf OnWebResourceRequested
    End Sub

    ''' <summary>
    ''' Enables offline mode (blocking http/https requests inside this WebView2 instance) if configured.
    ''' </summary>
    Private Sub EnableOfflineMode()
        If Not _disableInternetAfterLoad Then Return
        If _offlineModeEnabled Then Return
        _offlineModeEnabled = True
    End Sub

    ''' <summary>
    ''' WebView2 handler that blocks http/https requests when offline mode is enabled.
    ''' </summary>
    Private Sub OnWebResourceRequested(sender As Object, e As CoreWebView2WebResourceRequestedEventArgs)
        Try
            If Not _offlineModeEnabled Then Return

            Dim uri As Uri = Nothing
            If Not Uri.TryCreate(e.Request.Uri, UriKind.Absolute, uri) Then Return

            If uri.Scheme.Equals("http", StringComparison.OrdinalIgnoreCase) OrElse
               uri.Scheme.Equals("https", StringComparison.OrdinalIgnoreCase) Then

                Dim headers As String = "Content-Type: text/plain" & vbCrLf
                Dim stream As New MemoryStream(System.Text.Encoding.UTF8.GetBytes("Blocked by host (offline mode)."))

                e.Response = webView.CoreWebView2.Environment.CreateWebResourceResponse(
                    stream,
                    403,
                    "Forbidden",
                    headers)
            End If
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Requests an export of the current diagram via the diagrams.net embed postMessage API.
    ''' </summary>
    ''' <param name="format">The export format, e.g. pdf, png, svg, jpg.</param>
    Public Sub ExportToDevice(format As String)
        Try
            If webView?.CoreWebView2 Is Nothing Then Return

            Dim fmt As String = If(format, "").Trim().ToLowerInvariant()
            If String.IsNullOrWhiteSpace(fmt) Then Return

            If Not _hostIsReady Then
                _queuedExportFormat = fmt
                Return
            End If

            If _offlineModeEnabled Then
                MessageBox.Show(
                    "Export requires internet access. You opened draw.io in offline-after-load mode.",
                    $"{AN} - Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information)
                Return
            End If

            _pendingExportFormat = fmt

            Dim payload As New Dictionary(Of String, Object) From {
                {"action", "ri_export"},
                {"format", fmt}
            }

            Dim json As String = Newtonsoft.Json.JsonConvert.SerializeObject(payload)
            Dim script As String =
                $"window.__riRequestExport && window.__riRequestExport({Newtonsoft.Json.JsonConvert.SerializeObject(json)});"

            webView.CoreWebView2.ExecuteScriptAsync(script)

        Catch ex As Exception
            MessageBox.Show(
                $"Export failed: {ex.Message}",
                $"{AN} - Export",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Convenience method to export the current diagram as PDF.
    ''' </summary>
    Public Sub ExportPdfToDevice()
        ExportToDevice("pdf")
    End Sub

    ''' <summary>
    ''' Handles an export response message from the host HTML and persists the exported bytes to disk.
    ''' </summary>
    ''' <param name="message">The parsed JSON message payload from the host HTML.</param>
    Private Sub HandleExportFromJs(message As Dictionary(Of String, Object))
        Try
            Dim fmt As String = Nothing
            If message.ContainsKey("format") Then fmt = TryCast(message("format"), String)
            If String.IsNullOrWhiteSpace(fmt) Then fmt = If(_pendingExportFormat, "bin")

            Dim dataB64 As String = Nothing
            If message.ContainsKey("data") Then dataB64 = TryCast(message("data"), String)

            If String.IsNullOrWhiteSpace(dataB64) Then
                MessageBox.Show(
                    "Export did not return file data.",
                    $"{AN} - Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End If

            Dim suggestedFileName As String = Nothing
            If message.ContainsKey("suggestedFileName") Then suggestedFileName = TryCast(message("suggestedFileName"), String)
            If String.IsNullOrWhiteSpace(suggestedFileName) Then suggestedFileName = $"{_fileName}.{fmt}"

            Dim bytes As Byte()
            Try
                bytes = Convert.FromBase64String(dataB64)
            Catch
                MessageBox.Show(
                    "Export returned invalid base64 data.",
                    $"{AN} - Export",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning)
                Return
            End Try

            Me.Invoke(Sub()
                          Using sfd As New SaveFileDialog()
                              sfd.Title = "Save export"
                              sfd.FileName = suggestedFileName
                              sfd.Filter = "All files (*.*)|*.*"
                              Try
                                  sfd.InitialDirectory = Path.GetDirectoryName(_saveFilePath)
                              Catch
                              End Try

                              If sfd.ShowDialog(Me) <> DialogResult.OK Then Return
                              File.WriteAllBytes(sfd.FileName, bytes)
                          End Using
                      End Sub)

        Catch ex As Exception
            MessageBox.Show(
                $"Export failed: {ex.Message}",
                $"{AN} - Export",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' Builds the local HTML host that embeds diagrams.net and bridges messages between draw.io and the .NET host.
    ''' </summary>
    ''' <param name="xml">The mxfile XML payload.</param>
    ''' <param name="title">The diagram title.</param>
    ''' <returns>HTML document content.</returns>
    Private Function BuildHostHtml(xml As String, title As String) As String
        Dim xmlJs As String = Newtonsoft.Json.JsonConvert.SerializeObject(xml)
        Dim titleJs As String = Newtonsoft.Json.JsonConvert.SerializeObject(title)

        Return $"<!doctype html>
<html>
<head>
  <meta charset=""utf-8"">
  <title>{AN} - draw.io</title>
  <style>
    html, body {{ margin:0; padding:0; height:100%; overflow:hidden; }}
    #frame {{ width:100%; height:100%; border:0; }}
  </style>
</head>
<body>
  <iframe id=""frame"" src=""{DRAWIO_EMBED_URL}"" allow=""clipboard-read; clipboard-write; popups""></iframe>

  <script>
    (function() {{
      const frame = document.getElementById('frame');
      const xml = {xmlJs};
      const title = {titleJs};

      function postToHost(obj) {{
        try {{
          if (window.chrome && window.chrome.webview) {{
            window.chrome.webview.postMessage(JSON.stringify(obj));
          }}
        }} catch (e) {{}}
      }}

      function sendToDrawio(obj) {{
        try {{
          frame.contentWindow.postMessage(JSON.stringify(obj), '*');
          return true;
        }} catch (e) {{
          return false;
        }}
      }}

      window.__riRequestExport = function(jsonText) {{
        try {{
          const req = JSON.parse(jsonText);
          if (!req || req.action !== 'ri_export') return false;

          const fmt = (req.format || 'pdf');
          const exportMsg = {{
            action: 'export',
            format: fmt
          }};

          return sendToDrawio(exportMsg);
        }} catch (e) {{
          return false;
        }}
      }};

      window.addEventListener('message', function(evt) {{
        if (!evt || !evt.data || typeof evt.data !== 'string') return;

        let msg = null;
        try {{ msg = JSON.parse(evt.data); }} catch (e) {{ return; }}

        if (msg.event === 'init') {{
          sendToDrawio({{
            action: 'load',
            xml: xml,
            title: title
          }});
          postToHost({{ event: 'ri_loaded_sent' }});
          return;
        }}

        if (msg.event === 'export') {{
          postToHost({{
            event: 'ri_export',
            format: msg.format || '',
            data: msg.data || '',
            mime: msg.mime || '',
            suggestedFileName: (title || 'diagram') + '.' + (msg.format || 'bin')
          }});
          return;
        }}

        if (msg.event === 'save' || msg.event === 'exit') {{
          postToHost(msg);
          return;
        }}
      }}, false);

      postToHost({{ event: 'ri_host_ready' }});
    }})();
  </script>
</body>
</html>"
    End Function

    ''' <summary>
    ''' Handles the initial navigation completion. Used as a conservative fallback to enable offline mode once.
    ''' </summary>
    Private Sub OnNavigationCompleted(sender As Object, e As CoreWebView2NavigationCompletedEventArgs)
        If Not e.IsSuccess Then Return

        If Not _navigationCompletedOnce Then
            _navigationCompletedOnce = True

            ' Fallback only: the preferred point to enable offline mode is when we receive "ri_loaded_sent".
            If _disableInternetAfterLoad Then
                System.Threading.Tasks.Task.Delay(2000).ContinueWith(
                    Sub()
                        Try
                            If Not Me.IsDisposed Then EnableOfflineMode()
                        Catch
                        End Try
                    End Sub)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Receives messages from the host HTML and coordinates save/export and offline gating.
    ''' </summary>
    Private Sub OnWebMessageReceived(sender As Object, e As CoreWebView2WebMessageReceivedEventArgs)
        Try
            Dim messageString As String = e.TryGetWebMessageAsString()
            If String.IsNullOrEmpty(messageString) Then Return

            Dim message As Dictionary(Of String, Object) = Nothing
            Try
                message = Newtonsoft.Json.JsonConvert.DeserializeObject(Of Dictionary(Of String, Object))(messageString)
            Catch
                message = Nothing
            End Try
            If message Is Nothing Then Return

            Dim evtName As String = Nothing
            If message.ContainsKey("event") Then evtName = TryCast(message("event"), String)
            If String.IsNullOrEmpty(evtName) Then Return

            Select Case evtName
                Case "ri_host_ready"
                    _hostIsReady = True

                    If Not _hostReadyRaised Then
                        _hostReadyRaised = True
                        RaiseEvent HostReady(Me, EventArgs.Empty)
                    End If

                    If Not String.IsNullOrWhiteSpace(_queuedExportFormat) Then
                        Dim fmt As String = _queuedExportFormat
                        _queuedExportFormat = Nothing
                        ExportToDevice(fmt)
                    End If

                Case "ri_loaded_sent"
                    EnableOfflineMode()

                Case "save"
                    If message.ContainsKey("xml") Then
                        Dim savedXml As String = TryCast(message("xml"), String)
                        If Not String.IsNullOrWhiteSpace(savedXml) Then
                            SaveDiagram(savedXml)
                        End If
                    End If

                Case "exit"
                    Me.Close()

                Case "ri_export"
                    HandleExportFromJs(message)

            End Select

        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Saves the diagram XML to the configured file path and updates the window title temporarily.
    ''' </summary>
    ''' <param name="xmlContent">The mxfile XML to persist.</param>
    Private Sub SaveDiagram(xmlContent As String)
        Try
            File.WriteAllText(_saveFilePath, xmlContent, System.Text.Encoding.UTF8)

            Me.Text = $"{AN} - Draw.io Diagram Editor (Saved)"
            System.Threading.Tasks.Task.Delay(2000).ContinueWith(
                Sub()
                    If Not Me.IsDisposed Then
                        Try
                            Me.Invoke(Sub() Me.Text = $"{AN} - Draw.io Diagram Editor")
                        Catch
                        End Try
                    End If
                End Sub)

        Catch ex As Exception
            MessageBox.Show(
                $"Could not save the diagram: {ex.Message}",
                $"{AN} - Save Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' Disposes the hosted WebView2 when the form is closing.
    ''' </summary>
    Protected Overrides Sub OnFormClosing(e As FormClosingEventArgs)
        Try
            webView?.Dispose()
        Catch
        End Try

        MyBase.OnFormClosing(e)
    End Sub

End Class