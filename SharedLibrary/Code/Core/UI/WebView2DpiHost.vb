' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Shared native DPI contract for visible WinForms WebView2 hosts in Office.
' Do not change Office process awareness, HTML zoom or controller rasterization.
Option Strict On
Option Explicit On

Namespace SharedLibrary
    Public NotInheritable Class WebView2DpiHost
        Private Sub New()
        End Sub

        <System.Runtime.InteropServices.DllImport("user32.dll", SetLastError:=True)>
        Private Shared Function SetThreadDpiAwarenessContext(context As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetWindowDpiAwarenessContext(hwnd As System.IntPtr) As System.IntPtr
        End Function

        <System.Runtime.InteropServices.DllImport("user32.dll")>
        Private Shared Function GetDpiForWindow(hwnd As System.IntPtr) As System.UInt32
        End Function

        ' Only for synchronous HWND creation of independent top-level windows.
        ' Embedded controls must use their native parent's awareness instead.
        Public Shared Function StandaloneWindow() As System.IDisposable
            Return New WindowDpiScope(New System.IntPtr(-4)) ' Per Monitor v2
        End Function

        Public Shared Function MatchWindow(control As System.Windows.Forms.Control) As System.IDisposable
            Dim context As System.IntPtr = System.IntPtr.Zero
            Try
                If control IsNot Nothing AndAlso Not control.IsDisposed Then
                    If control.IsHandleCreated Then
                        context = GetWindowDpiAwarenessContext(control.Handle)
                    ElseIf control.Parent IsNot Nothing AndAlso control.Parent.IsHandleCreated Then
                        context = GetWindowDpiAwarenessContext(control.Parent.Handle)
                    End If
                End If
            Catch ex As System.EntryPointNotFoundException
            Catch ex As System.DllNotFoundException
            End Try
            Return New WindowDpiScope(context)
        End Function

        Public Shared Function GetWindowDpi(hwnd As System.IntPtr) As System.UInt32
            Try
                Return GetDpiForWindow(hwnd)
            Catch ex As System.EntryPointNotFoundException
                Return 0UI
            Catch ex As System.DllNotFoundException
                Return 0UI
            End Try
        End Function

        Public Shared Sub SynchronizeLayout(webView As Microsoft.Web.WebView2.WinForms.WebView2)
            If webView Is Nothing OrElse webView.IsDisposed Then Return
            Using scope As System.IDisposable = MatchWindow(webView)
                ' Preserve toolbars, panels and each existing Dock/Anchor arrangement.
                If webView.Parent IsNot Nothing Then webView.Parent.PerformLayout()
                webView.PerformLayout()
            End Using
        End Sub

        ' Start controller creation under the actual HWND's context, then restore
        ' before yielding. A thread DPI scope must NEVER survive an Await.
        Public Shared Async Function EnsureInitializedAsync(webView As Microsoft.Web.WebView2.WinForms.WebView2,
                environment As Microsoft.Web.WebView2.Core.CoreWebView2Environment) As System.Threading.Tasks.Task
            If webView Is Nothing Then Throw New System.ArgumentNullException(NameOf(webView))
            If webView.IsDisposed Then Throw New System.ObjectDisposedException(NameOf(webView))
            If Not webView.IsHandleCreated Then webView.CreateControl()
            Dim initialization As System.Threading.Tasks.Task
            Using scope As System.IDisposable = MatchWindow(webView)
                SynchronizeLayout(webView)
                initialization = webView.EnsureCoreWebView2Async(environment)
            End Using
            Await initialization
            If webView.IsDisposed Then Throw New System.ObjectDisposedException(NameOf(webView))
            SynchronizeLayout(webView)
        End Function

        Private NotInheritable Class WindowDpiScope
            Implements System.IDisposable
            Private _previous As System.IntPtr
    
            Public Sub New(context As System.IntPtr)
                If context = System.IntPtr.Zero Then Return
                Try
                    _previous = SetThreadDpiAwarenessContext(context)
                    If _previous = System.IntPtr.Zero Then System.Diagnostics.Debug.WriteLine("[WebView2 DPI] Context switch failed; using inherited context. Win32 error=" & System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
                Catch ex As System.EntryPointNotFoundException
                    System.Diagnostics.Debug.WriteLine("[WebView2 DPI] Thread DPI contexts are unavailable; using inherited context.")
                Catch ex As System.DllNotFoundException
                    System.Diagnostics.Debug.WriteLine("[WebView2 DPI] Thread DPI contexts are unavailable; using inherited context.")
                End Try
            End Sub
    
            Public Sub Dispose() Implements System.IDisposable.Dispose
                If _previous = System.IntPtr.Zero Then Return
                Dim previous As System.IntPtr = _previous
                _previous = System.IntPtr.Zero
                If SetThreadDpiAwarenessContext(previous) = System.IntPtr.Zero Then
                    System.Diagnostics.Debug.WriteLine("[WebView2 DPI] Could not restore caller context. Win32 error=" & System.Runtime.InteropServices.Marshal.GetLastWin32Error().ToString(System.Globalization.CultureInfo.InvariantCulture))
                End If
            End Sub
        End Class
    End Class

    ' HWND creation must happen in the intended context even when Office invokes
    ' an add-in on a System-aware thread. Ownership is not child-window parenting.
    Public Class WebView2HostForm
        Inherits System.Windows.Forms.Form

        Public Sub New()
            AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        End Sub

        Protected Overrides Sub CreateHandle()
            If TopLevel AndAlso Parent Is Nothing Then
                Using scope As System.IDisposable = WebView2DpiHost.StandaloneWindow()
                    MyBase.CreateHandle()
                End Using
            Else
                Using scope As System.IDisposable = WebView2DpiHost.MatchWindow(Parent)
                    MyBase.CreateHandle()
                End Using
            End If
        End Sub

        <System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)>
        Private Structure WindowRectangle
            Public Left As System.Int32
            Public Top As System.Int32
            Public Right As System.Int32
            Public Bottom As System.Int32
        End Structure

        Protected Overrides Sub WndProc(ByRef message As System.Windows.Forms.Message)
            Const WmDpiChanged As System.Int32 = &H2E0
            Dim suggested As System.Nullable(Of System.Drawing.Rectangle) = Nothing
            If TopLevel AndAlso message.Msg = WmDpiChanged AndAlso message.LParam <> System.IntPtr.Zero Then
                Dim rect As WindowRectangle = System.Runtime.InteropServices.Marshal.PtrToStructure(Of WindowRectangle)(message.LParam)
                If rect.Right > rect.Left AndAlso rect.Bottom > rect.Top Then
                    suggested = System.Drawing.Rectangle.FromLTRB(rect.Left, rect.Top, rect.Right, rect.Bottom)
                End If
            End If
            Using scope As System.IDisposable = WebView2DpiHost.MatchWindow(Me)
                MyBase.WndProc(message)
                If suggested.HasValue AndAlso Not IsDisposed Then
                    ' .NET 4.8 inherits Office configuration; honour the native
                    ' suggested rectangle even when automatic resizing is disabled.
                    If Bounds <> suggested.Value Then Bounds = suggested.Value
                    PerformLayout()
                End If
            End Using
        End Sub
    End Class

    ' Native WebView bounds updates also occur after Office callbacks / layout.
    ' Match the existing window, never force PMv2 on an Office-owned child HWND.
    Public Class DpiAwareWebView2
        Inherits Microsoft.Web.WebView2.WinForms.WebView2

        Protected Overrides Sub CreateHandle()
            Using scope As System.IDisposable = WebView2DpiHost.MatchWindow(Parent)
                MyBase.CreateHandle()
            End Using
        End Sub

        Protected Overrides Sub OnSizeChanged(e As System.EventArgs)
            Using scope As System.IDisposable = WebView2DpiHost.MatchWindow(Me)
                MyBase.OnSizeChanged(e)
            End Using
        End Sub

        Protected Overrides Sub OnLocationChanged(e As System.EventArgs)
            Using scope As System.IDisposable = WebView2DpiHost.MatchWindow(Me)
                MyBase.OnLocationChanged(e)
            End Using
        End Sub
    End Class
End Namespace
