Option Strict On
Option Explicit On
Namespace SharedLibrary
    Public Partial Class SharedMethods
        Public Const AN As String = "Red Ink regression probe"
        Public Enum LogoType
            Standard
        End Enum
        Private Shared ReadOnly Logo As New System.Drawing.Bitmap(20, 20)
        Public Shared Function GetLogoBitmap(kind As LogoType) As System.Drawing.Bitmap
            Return Logo
        End Function
        Public Shared Function GetHostWindowHandle() As System.IntPtr
            Return System.IntPtr.Zero
        End Function
        Public Shared Function IsWindowOnCurrentThread(value As System.IntPtr) As Boolean
            Return False
        End Function
        Public Shared Function IfOwnerOnCurrentThread(owner As System.Windows.Forms.IWin32Window) As System.Windows.Forms.IWin32Window
            Return Nothing
        End Function
        Public Shared Sub PutInClipboard(text As String)
            System.Windows.Forms.Clipboard.SetText(text)
        End Sub
        Public Shared Function ResolveSameThreadDialogOwner() As System.Windows.Forms.IWin32Window
            Return Nothing
        End Function
        Public Shared Function PushDialogOwner(owner As System.Windows.Forms.IWin32Window) As System.IDisposable
            Return New NullScope()
        End Function
        Public Shared Sub ForceDialogToForeground(form As System.Windows.Forms.Form)
        End Sub
        Public Shared Sub AttachForeignForegroundWatchdog(form As System.Windows.Forms.Form)
        End Sub
        Public Shared Sub ShowCustomMessageBox(text As String)
            System.Windows.Forms.MessageBox.Show(text)
        End Sub
        Public Shared Function CreateMarkdownHtmlPipeline(Optional useSoftlineBreakAsHardlineBreak As Boolean = False, Optional usePreciseSourceLocation As Boolean = False) As Markdig.MarkdownPipeline
            Return New Markdig.MarkdownPipelineBuilder().Build()
        End Function
        Public Shared Function NormalizeMarkdownForHtmlDisplay(text As String) As String
            Return text
        End Function
        Public Shared Function CreateOcrStatusForProbe() As System.IDisposable
            Return New OcrChunkStatusDialog()
        End Function
        Public Shared Function StartOcrStatusForProbe(status As System.IDisposable, token As System.Threading.CancellationToken) As System.Threading.Tasks.Task
            Return DirectCast(status, OcrChunkStatusDialog).ShowAsync("Probe OCR status", token)
        End Function
        Public Shared Function OcrCompletionForProbe(status As System.IDisposable) As System.Threading.Tasks.Task
            Return DirectCast(status, OcrChunkStatusDialog).Completion
        End Function
        Private Class NullScope
            Implements System.IDisposable
            Public Sub Dispose() Implements System.IDisposable.Dispose
            End Sub
        End Class
    End Class
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private ReadOnly value As System.IntPtr
        Public Sub New(handle As System.IntPtr)
            value = handle
        End Sub
        Public ReadOnly Property Handle As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return value
            End Get
        End Property
    End Class
    Public Class RiCrashLogger
        Public Shared ReadOnly Failures As New System.Collections.Concurrent.ConcurrentQueue(Of String)()
        Public Shared Sub LogMarker(name As String, details As String)
            System.Diagnostics.Debug.WriteLine(name & ": " & details)
        End Sub
        Public Shared Sub LogException(name As String, errorValue As Exception)
            Failures.Enqueue(name & ": " & errorValue.ToString())
        End Sub
    End Class
    Public Class OfficeWindowWatchdog
        Public Shared Sub InspectDialogOwner(owner As System.Windows.Forms.IWin32Window, name As String, caller As String)
        End Sub
    End Class
End Namespace
