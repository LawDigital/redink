' Part of "Red Ink" (Ask Inky)
' Copyright (c) LawDigital AG, Switzerland. All rights reserved. For license to use see https://redink.ai.
' =============================================================================
' File: Program.vb
' Purpose: Main Windows entry point and application initialization.
' Version: V121026 (12 October 2026). Update on each dated delivery.
' Architecture: Windows native host with embedded Office Web-based interface.
' Notes: Red Ink License and Ask Inky third-party attributions are embedded in LICENSE.txt.
' =============================================================================
' Third-party elements: Markdig (BSD-2-Clause), Microsoft.Web.WebView2 (Microsoft license),
' .NET runtime components (applicable Microsoft/.NET licenses), adapted Lucide/Feather SVG paths
' (ISC/MIT) and adapted Office Web styling; full notices: LICENSE.txt.
'
' Part of Red Ink (AskInky). Standalone WinForms message loop, not a VSTO add-in.
Friend Module Program
    Public Const Version As String = "V221126"
    ' 2.0.4: a second start (Start menu, installer) asks the running Ask Inky to open itself.
    Friend Const ShowSignalName As String = "Local\RedInk.AskInky.Show"
    Friend Const AutostartArgument As String = "--autostart"
    Friend ReadOnly PendingConnectPath As String = System.IO.Path.Combine(StoragePaths.Root, "pending-connect.txt")
    <System.STAThread>
    Public Sub Main(args As String())
        AddHandler System.AppDomain.CurrentDomain.UnhandledException, AddressOf LogUnhandledException
        ' P18 W367 (DARO UID 551): "Ask Inky mit diesem Server verbinden" on the Helper's download page opens
        ' redink-askinky://connect?helper=<address>. Only the address is passed, never a credential; the
        ' running instance asks the person before it uses it.
        Dim connect As String = ConnectTarget(args)
        Dim created As Boolean
        Using instance As New System.Threading.Mutex(True, "Local\RedInk.AskInky.Instance", created)
            If connect.Length > 0 Then
                Try
                    System.IO.Directory.CreateDirectory(StoragePaths.Root)
                    System.IO.File.WriteAllText(PendingConnectPath, connect, New System.Text.UTF8Encoding(False))
                Catch ex As System.Exception
                    ClientLog.Write("connect.pending.write.failed", ex)
                End Try
            End If
            Dim autostart As Boolean = args IsNot Nothing AndAlso System.Array.Exists(args, Function(a) String.Equals(a, AutostartArgument, System.StringComparison.OrdinalIgnoreCase))
            If Not created Then
                ' One Ask Inky per Windows session; the running one picks up the address, or opens itself.
                If connect.Length = 0 AndAlso Not autostart Then
                    Try
                        Using signal As New System.Threading.EventWaitHandle(False, System.Threading.EventResetMode.AutoReset, ShowSignalName)
                            signal.Set()
                        End Using
                    Catch ex As System.Exception
                        ClientLog.Write("instance.signal.send.failed", ex)
                    End Try
                End If
                Return
            End If
            System.Windows.Forms.Application.SetHighDpiMode(System.Windows.Forms.HighDpiMode.PerMonitorV2)
            System.Windows.Forms.Application.EnableVisualStyles()
            System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
            System.Windows.Forms.Application.Run(New AssistantContext(autostart))
        End Using
    End Sub

    ' redink-askinky://connect?helper=https://server  ->  normalized https address, or "".
    Friend Function ConnectTarget(args As String()) As String
        If args Is Nothing OrElse args.Length = 0 Then Return ""
        Dim uri As System.Uri = Nothing
        Dim raw As String = If(args(0), "").Trim()
        If raw.Length > 2048 OrElse Not System.Uri.TryCreate(raw, System.UriKind.Absolute, uri) OrElse
           Not String.Equals(uri.Scheme, "redink-askinky", System.StringComparison.OrdinalIgnoreCase) OrElse
           Not String.Equals(uri.Host, "connect", System.StringComparison.OrdinalIgnoreCase) Then Return ""
        For Each pair As String In uri.Query.TrimStart("?"c).Split("&"c)
            Dim kv As String() = pair.Split("="c, 2)
            If kv.Length = 2 AndAlso kv(0) = "helper" Then
                Try
                    Dim value As String = HelperClient.NormalizeHelperUrl(System.Uri.UnescapeDataString(kv(1)))
                    If value.StartsWith("https://", System.StringComparison.OrdinalIgnoreCase) Then Return value
                Catch
                    Return ""
                End Try
            End If
        Next
        Return ""
    End Function

    Private Sub LogUnhandledException(sender As Object, e As System.UnhandledExceptionEventArgs)
        ClientLog.Write("application.unhandled." & If(e.IsTerminating, "terminating", "reported"), TryCast(e.ExceptionObject, System.Exception))
    End Sub
End Module
