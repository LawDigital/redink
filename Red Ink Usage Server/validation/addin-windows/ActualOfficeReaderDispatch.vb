Option Strict On
Option Explicit On
Namespace Global.NativeOfficeReaderProbe
Public Class OfficeReaderDispatchProbe
    Public LastReaderThreadId As System.Int32
    Public Function ReadOfficeForProbe(path As System.String) As System.Threading.Tasks.Task(Of System.Tuple(Of System.Boolean, System.String, System.String))
        Return TryExtractOfficeTextAsync(path)
    End Function
    Public Function ReadTextForProbe(path As System.String) As System.Threading.Tasks.Task(Of System.Tuple(Of System.Boolean, System.String, System.String))
        Return TryExtractTextLikeAsync(path)
    End Function
    ' Only the I/O and COM reader bodies are test doubles. The routing methods below
    ' are extracted byte-for-text from production and use the actual host dispatcher.
    Private Function TryExtractOfficeText(path As System.String, ByRef text As System.String, ByRef label As System.String) As System.Boolean
        LastReaderThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId
        text = "office:" & path
        label = "office label"
        Return True
    End Function
    Private Function TryExtractTextLike(path As System.String, ByRef text As System.String, ByRef label As System.String) As System.Boolean
        LastReaderThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId
        text = "text:" & path
        label = "text label"
        Return True
    End Function
    Private Async Function TryExtractOfficeTextAsync(filePath As System.String) As System.Threading.Tasks.Task(Of System.Tuple(Of System.Boolean, System.String, System.String))
        Dim reader As System.Func(Of System.Tuple(Of System.Boolean, System.String, System.String)) =
            Function()
                ' Outputs are local to the dispatched operation; no ByRef parameter crosses threads.
                Dim extracted As System.String = Nothing
                Dim label As System.String = Nothing
                Dim succeeded As System.Boolean = TryExtractOfficeText(filePath, extracted, label)
                Return System.Tuple.Create(succeeded, extracted, label)
            End Function
        Dim extension As System.String = If(System.String.IsNullOrWhiteSpace(filePath), System.String.Empty,
                                             System.IO.Path.GetExtension(filePath).ToLowerInvariant())
        Select Case extension
            Case ".doc", ".rtf", ".xls", ".ppt"
                ' Keep the complete COM create/read/close/release sequence on the actual host STA.
                Return Await SharedLibrary.SharedLibrary.SharedMethods.PostToHostUiAsync(reader).ConfigureAwait(False)
            Case Else
                ' Sandboxed OpenXML/mail readers and ordinary file I/O remain on the caller.
                Return reader.Invoke()
        End Select
    End Function
    Private Async Function TryExtractTextLikeAsync(filePath As System.String) As System.Threading.Tasks.Task(Of System.Tuple(Of System.Boolean, System.String, System.String))
        Dim reader As System.Func(Of System.Tuple(Of System.Boolean, System.String, System.String)) =
            Function()
                Dim extracted As System.String = Nothing
                Dim label As System.String = Nothing
                Dim succeeded As System.Boolean = TryExtractTextLike(filePath, extracted, label)
                Return System.Tuple.Create(succeeded, extracted, label)
            End Function
        If Not System.String.IsNullOrWhiteSpace(filePath) AndAlso
           System.String.Equals(System.IO.Path.GetExtension(filePath), ".rtf", System.StringComparison.OrdinalIgnoreCase) Then
            ' The existing RTF fallback uses Word; its entire COM lifetime belongs to the host.
            Return Await SharedLibrary.SharedLibrary.SharedMethods.PostToHostUiAsync(reader).ConfigureAwait(False)
        End If
        Return reader.Invoke()
    End Function
End Class
End Namespace
