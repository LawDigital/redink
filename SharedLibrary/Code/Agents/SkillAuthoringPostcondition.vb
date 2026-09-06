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
        Private Shared ReadOnly _resourceRootWrite As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>Workflow ids that wrote a skill/agent-like file into the temporary workspace.</summary>
        Private Shared ReadOnly _workspaceSkillWrite As New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)

        ''' <summary>Exact successful resource-root write paths, grouped by workflow.</summary>
        Private Shared ReadOnly _resourceRootPaths As New System.Collections.Generic.Dictionary(Of System.String, System.Collections.Generic.HashSet(Of System.String))(System.StringComparer.OrdinalIgnoreCase)

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
        Public Shared Sub NoteResourceRootWrite(Optional fullPath As System.String = Nothing)
            Dim key As System.String = CurrentKey()
            If key.Length = 0 Then Return
            SyncLock _sync
                _resourceRootWrite.Add(key)
                If Not System.String.IsNullOrWhiteSpace(fullPath) Then
                    Dim paths As System.Collections.Generic.HashSet(Of System.String) = Nothing
                    If Not _resourceRootPaths.TryGetValue(key, paths) Then
                        paths = New System.Collections.Generic.HashSet(Of System.String)(System.StringComparer.OrdinalIgnoreCase)
                        _resourceRootPaths(key) = paths
                    End If
                    paths.Add(System.IO.Path.GetFullPath(fullPath))
                End If
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

        ''' <summary>
        ''' Validates every authored SKILL.md/AGENT.md touched in the current resource root. Reference-only
        ''' mutations are allowed; when a descriptor was touched, its frontmatter must satisfy the runtime loader.
        ''' </summary>
        Public Shared Function HasValidAuthoredResourceStructure(ByRef failureReason As System.String,
                                                                 Optional context As System.Object = Nothing) As System.Boolean
            failureReason = System.String.Empty
            Dim key As System.String = CurrentKey()
            If key.Length = 0 Then Return True

            Dim candidates As New System.Collections.Generic.List(Of System.String)()
            SyncLock _sync
                Dim paths As System.Collections.Generic.HashSet(Of System.String) = Nothing
                If Not _resourceRootPaths.TryGetValue(key, paths) OrElse paths Is Nothing Then Return True

                For Each fullPath As System.String In paths
                    If IsAuthoredDescriptorPath(fullPath) Then
                        candidates.Add(fullPath)
                    End If
                Next
            End SyncLock

            For Each candidate As System.String In candidates
                Dim detail As System.String = System.String.Empty
                If Not AgentResources.TryValidateAuthoredResourceFrontmatter(candidate, detail) Then
                    failureReason = System.IO.Path.GetFileName(candidate) & ": " & detail
                    Return False
                End If
            Next

            Return True
        End Function


        Private Shared Function IsAuthoredDescriptorPath(fullPath As System.String) As System.Boolean
            If System.String.IsNullOrWhiteSpace(fullPath) Then Return False

            Dim leaf As System.String = System.IO.Path.GetFileName(fullPath)
            If System.String.Equals(leaf, "SKILL.md", System.StringComparison.OrdinalIgnoreCase) OrElse
               System.String.Equals(leaf, "AGENT.md", System.StringComparison.OrdinalIgnoreCase) Then
                Return True
            End If

            ' Flat agent descriptors under ...\agents\<name>.md are supported by the runtime loader.
            If Not System.String.Equals(System.IO.Path.GetExtension(leaf), ".md", System.StringComparison.OrdinalIgnoreCase) Then
                Return False
            End If

            Dim parent As System.String = System.IO.Path.GetDirectoryName(fullPath)
            If System.String.IsNullOrWhiteSpace(parent) Then Return False

            Return System.String.Equals(
                System.IO.Path.GetFileName(parent),
                "agents",
                System.StringComparison.OrdinalIgnoreCase)
        End Function

        Public Shared Function BuildStructureGuardPrompt(failureReason As System.String) As System.String
            Return "The authored Skill/Agent descriptor exists under the resource root but its runtime frontmatter is invalid (" &
                   If(failureReason, "invalid frontmatter") & "). Repair the SAME resource file in place. " &
                   "It must start with YAML frontmatter delimited by --- and include non-empty name and description fields plus allowed-tools. " &
                   "After writing, re-read/validate the descriptor before reporting completion."
        End Function

        ''' <summary>Clears recorded state for a finished workflow (best-effort).</summary>
        Public Shared Sub Clear(workflowId As String)
            If String.IsNullOrWhiteSpace(workflowId) Then Return
            SyncLock _sync
                _resourceRootWrite.Remove(workflowId)
                _workspaceSkillWrite.Remove(workflowId)
                _resourceRootPaths.Remove(workflowId)
            End SyncLock
        End Sub

    End Class

End Namespace
