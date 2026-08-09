' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SkillAuthoringPostcondition.vb
' Purpose: Workflow-scoped postcondition for skill/agent authoring tasks. Records,
'          per workflow, whether a write landed under an authorized resource root
'          versus a skill/agent-like file written into the temporary workspace.
'          The tooling loop consults this before accepting a 'complete' final turn
'          so a skill/agent authoring task cannot be satisfied by a workspace_write.
'
' Capability-driven: classification is based on the resolved write target (resource
' root vs. workspace) and structural markers (SKILL.md/AGENT.md, skills/ or agents/
' path segments) - never on request text or tool-name heuristics.
' =============================================================================

Option Strict On
Option Explicit On

Namespace Agents

    Public NotInheritable Class SkillAuthoringPostcondition

        Private Sub New()
        End Sub

        Private Shared ReadOnly _sync As New Object()

        ''' <summary>Workflow ids that produced at least one successful write under a resource root.</summary>
        Private Shared ReadOnly _resourceRootWrite As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Workflow ids that wrote a skill/agent-like file into the temporary workspace.</summary>
        Private Shared ReadOnly _workspaceSkillWrite As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)

        ''' <summary>Guidance injected when the postcondition rejects a 'complete' turn.</summary>
        Public Shared ReadOnly Property GuardPrompt As String
            Get
                Return "The task involves creating or modifying a Red Ink Skill or Agent, but the only " &
                       "output was written into the temporary workspace, which does NOT install a skill. " &
                       "Use the skill-author skill and the resource filesystem tools (file_make_dir, " &
                       "file_copy, text_write) to create the skill/agent under the resource root using an " &
                       "ABSOLUTE path (e.g. new_resource_root + '\skills\<name>\SKILL.md'), then finish."
            End Get
        End Property

        Private Shared Function CurrentKey() As String
            Dim key As String = WorkflowContinuity.CurrentWorkflowId
            Return If(String.IsNullOrWhiteSpace(key), "", key)
        End Function

        ''' <summary>Records a successful write that resolved under an authorized resource root.</summary>
        Public Shared Sub NoteResourceRootWrite()
            Dim key As String = CurrentKey()
            If key.Length = 0 Then Return
            SyncLock _sync
                _resourceRootWrite.Add(key)
            End SyncLock
        End Sub

        ''' <summary>Records a skill/agent-like file written into the temporary workspace (wrong target).</summary>
        Public Shared Sub NoteWorkspaceSkillLikeWrite()
            Dim key As String = CurrentKey()
            If key.Length = 0 Then Return
            SyncLock _sync
                _workspaceSkillWrite.Add(key)
            End SyncLock
        End Sub

        ''' <summary>
        ''' True when this run is a skill/agent authoring task that must mutate a resource root:
        ''' author mode is active AND a skill/agent-like file was written into the workspace.
        ''' </summary>
        Public Shared Function RequiresSkillRootMutation(Optional context As Object = Nothing) As Boolean
            If Not SkillAuthorMode.IsActive Then Return False
            Dim key As String = CurrentKey()
            If key.Length = 0 Then Return False
            SyncLock _sync
                Return _workspaceSkillWrite.Contains(key)
            End SyncLock
        End Function

        ''' <summary>True when this run wrote at least once under an authorized resource root.</summary>
        Public Shared Function HasSkillRootMutation(Optional context As Object = Nothing) As Boolean
            Dim key As String = CurrentKey()
            If key.Length = 0 Then Return False
            SyncLock _sync
                Return _resourceRootWrite.Contains(key)
            End SyncLock
        End Function

        ''' <summary>Clears recorded state for a finished workflow (best-effort).</summary>
        Public Shared Sub Clear(workflowId As String)
            If String.IsNullOrWhiteSpace(workflowId) Then Return
            SyncLock _sync
                _resourceRootWrite.Remove(workflowId)
                _workspaceSkillWrite.Remove(workflowId)
            End SyncLock
        End Sub

    End Class

End Namespace
