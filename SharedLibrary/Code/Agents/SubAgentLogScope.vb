Option Explicit On
Option Strict On

Namespace Agents
    ' Presentation-only async-flow identity. No host/model/registry or workflow state.
    ' Immutable frames ensure nested child calls cannot relabel their parent or other jobs.
    Public NotInheritable Class SubAgentLogScope
        Implements System.IDisposable

        Private Shared ReadOnly Current As New System.Threading.AsyncLocal(Of SubAgentLogScope)()
        Private ReadOnly _parent As SubAgentLogScope
        Private ReadOnly _path As System.String
        Private _disposed As System.Boolean

        Private Sub New(agentName As System.String)
            _parent = Current.Value
            Dim name As System.String = System.Text.RegularExpressions.Regex.Replace(
                If(agentName, System.String.Empty), "[\s\[\]]+", " ").Trim()
            If name.Length = 0 Then name = "unnamed"
            If name.Length > 80 Then name = name.Substring(0, 80)
            _path = If(_parent Is Nothing, name, _parent._path & " > " & name)
        End Sub

        Public Shared Function Enter(agentName As System.String) As SubAgentLogScope
            Dim scope As New SubAgentLogScope(agentName)
            Current.Value = scope
            Return scope
        End Function

        Public Shared Function FormatMessage(message As System.String, Optional subAgentFallback As System.Boolean = False) As System.String
            Dim text As System.String = If(message, System.String.Empty)
            Dim scope As SubAgentLogScope = Current.Value
            If scope Is Nothing Then Return If(subAgentFallback, "[Subagent] " & text, text)
            Dim prefix As System.String = "[Subagent " & scope._path & "] "
            If text.StartsWith(prefix, System.StringComparison.Ordinal) Then Return text
            Return prefix & text
        End Function

        Public Sub Dispose() Implements System.IDisposable.Dispose
            If _disposed Then Return
            _disposed = True
            If System.Object.ReferenceEquals(Current.Value, Me) Then Current.Value = _parent
        End Sub
    End Class
End Namespace
