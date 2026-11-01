Option Strict On
Option Explicit On
Namespace Agents
    Public NotInheritable Class AskUserOption
        Public Property Id As String
        Public Property Label As String
        Public Property Description As String
    End Class

    ''' <summary>Structured request passed to the host UI callback.</summary>
    Public NotInheritable Class AskUserRequest
        Public Property Question As String
        Public Property Options As List(Of AskUserOption)
        Public Property AllowFreeText As Boolean = True
        Public Property MultiSelect As Boolean = False
        ''' <summary>"text" (default), "integer", "number", or "choice".</summary>
        Public Property InputType As String = "text"
    End Class

    ''' <summary>Structured answer returned by the host UI callback.</summary>
    Public NotInheritable Class AskUserResult
        ''' <summary>"answered" or "cancelled".</summary>
        Public Property Status As String = "cancelled"
        Public Property SelectedOptionIds As List(Of String)
        Public Property FreeText As String
    End Class

End Namespace
