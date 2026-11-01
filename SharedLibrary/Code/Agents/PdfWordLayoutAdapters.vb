' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' Optional provider-agnostic PDF-to-Word layout-conversion adapter contract.

' =============================================================================
' File: PdfWordLayoutAdapters.vb
' Purpose:
'   Provider-neutral PDF-to-Word layout conversion request/result contracts and optional
'   adapter registry.
'
' Architecture / Function:
'   Adapters register and resolve through one shared boundary; the contract does not
'   select a provider-specific implementation.
' =============================================================================

Option Strict On
Option Explicit On
Option Infer On

Namespace Agents

    Public NotInheritable Class PdfWordLayoutConversionRequest
        Public Property InputPath As System.String = System.String.Empty
        Public Property OutputPath As System.String = System.String.Empty
        Public Property SourceSha256 As System.String = System.String.Empty
    End Class

    Public NotInheritable Class PdfWordLayoutConversionResult
        Public Property Success As System.Boolean
        Public Property AdapterId As System.String = System.String.Empty
        Public Property Quality As System.String = "unknown"
        Public Property OutputPath As System.String = System.String.Empty
        Public Property ErrorCode As System.String = System.String.Empty
        Public Property Message As System.String = System.String.Empty
    End Class

    Public Interface IPdfWordLayoutAdapter
        ReadOnly Property AdapterId As System.String
        ReadOnly Property AdapterVersion As System.String
        Function IsAvailable() As System.Boolean
        Function ConvertAsync(request As PdfWordLayoutConversionRequest,
                              cancellationToken As System.Threading.CancellationToken) As System.Threading.Tasks.Task(Of PdfWordLayoutConversionResult)
    End Interface

    Public NotInheritable Class PdfWordLayoutAdapterRegistry

        Private Shared ReadOnly SyncRoot As New System.Object()
        Private Shared ReadOnly Adapters As New System.Collections.Generic.Dictionary(Of System.String, IPdfWordLayoutAdapter)(System.StringComparer.OrdinalIgnoreCase)

        Private Sub New()
        End Sub

        Public Shared Sub Register(adapter As IPdfWordLayoutAdapter)
            If adapter Is Nothing Then Throw New System.ArgumentNullException(NameOf(adapter))
            Dim adapterId As System.String = If(adapter.AdapterId, System.String.Empty).Trim()
            If adapterId.Length = 0 Then Throw New System.ArgumentException("A layout adapter must expose a non-empty AdapterId.", NameOf(adapter))

            SyncLock SyncRoot
                Adapters(adapterId) = adapter
            End SyncLock
        End Sub

        Public Shared Function TryResolve(requestedAdapterId As System.String,
                                          ByRef adapter As IPdfWordLayoutAdapter) As System.Boolean
            adapter = Nothing
            SyncLock SyncRoot
                If Not System.String.IsNullOrWhiteSpace(requestedAdapterId) Then
                    Dim requested As IPdfWordLayoutAdapter = Nothing
                    If Adapters.TryGetValue(requestedAdapterId.Trim(), requested) AndAlso IsAdapterAvailable(requested) Then
                        adapter = requested
                        Return True
                    End If
                    Return False
                End If

                For Each candidate As IPdfWordLayoutAdapter In Adapters.Values
                    If IsAdapterAvailable(candidate) Then
                        adapter = candidate
                        Return True
                    End If
                Next
            End SyncLock
            Return False
        End Function

        Public Shared Function GetAvailableAdapterIds() As System.Collections.Generic.List(Of System.String)
            Dim result As New System.Collections.Generic.List(Of System.String)()
            SyncLock SyncRoot
                For Each pair As System.Collections.Generic.KeyValuePair(Of System.String, IPdfWordLayoutAdapter) In Adapters
                    If IsAdapterAvailable(pair.Value) Then result.Add(pair.Key)
                Next
            End SyncLock
            result.Sort(System.StringComparer.OrdinalIgnoreCase)
            Return result
        End Function

        Private Shared Function IsAdapterAvailable(adapter As IPdfWordLayoutAdapter) As System.Boolean
            If adapter Is Nothing Then Return False
            Try
                Return adapter.IsAvailable()
            Catch ex As System.Exception
                System.Diagnostics.Debug.WriteLine("PDF layout adapter availability check failed: " & ex.Message)
                Return False
            End Try
        End Function

    End Class

End Namespace
