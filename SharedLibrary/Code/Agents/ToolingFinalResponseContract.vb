' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ToolingFinalResponseContract.vb
' Purpose:
'   Defines the small shared contract that tells a tooling loop how a model's final
'   non-tool response must be interpreted and formatted.
'
' Architecture / Function:
'   - UserFacingTaskStatus requires the machine-readable TASK_STATUS completion footer.
'   - RawCallerText is used for nested/internal callers that must receive unwrapped text.
'   - Helper functions centralize contract serialization and footer requirements so
'     Word, Outlook and shared sub-agent flows do not diverge.
' =============================================================================


Option Explicit On
Option Strict On

Namespace Agents

    Public Enum ToolingFinalResponseContract
        UserFacingTaskStatus = 0
        RawCallerText = 1
    End Enum

    Public Module ToolingFinalResponseContractHelpers

        Public Function RequiresTaskStatusFooter(contract As ToolingFinalResponseContract) As Boolean
            Return contract = ToolingFinalResponseContract.UserFacingTaskStatus
        End Function

        Public Function FormatToolingFinalResponseContract(contract As ToolingFinalResponseContract) As String
            Select Case contract
                Case ToolingFinalResponseContract.RawCallerText
                    Return "raw_caller_text"
                Case Else
                    Return "user_facing_task_status"
            End Select
        End Function

    End Module

End Namespace