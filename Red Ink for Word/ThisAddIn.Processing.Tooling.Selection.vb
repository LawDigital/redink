' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Processing.Tooling.Selection.vb
' Purpose: Tool selection UI, persistence, and availability management.
'
' Responsibilities:
'  - Load tool configurations from INI files (external/connector tools).
'  - Register transport-backed tools in the executor registry.
'  - Show tool selection dialogs with main/advanced/workspace tabs.
'  - Persist selected tool names to application settings.
'  - Load persisted tool selections on startup.
'  - Segregate tools by type (main/basic vs. advanced/specialty).
'  - Classify advanced tools (skills, agents, text tools, workspace tools, js_run).
'  - Build effective tool list considering workspace connection state.
'  - Expose tool selection via "Discuss Inky" terminology and multi-tab UI.
'  - Support skills/agents/memory/workspace management dialogs.
'  - Handle legacy tool selection migration (old format -> new segmented format).
'
' Architecture:
'  - Main tools: external connectors, web/search/knowledge retrieval.
'  - Advanced tools: skills, agents, workspace access, scripting.
'  - Conditional display: workspace tools only when workspace connected.
'  - Settings keys: SelectedMainToolNames, SelectedAdvancedToolNames, AdvancedToolsEnabled.
'
' External Dependencies:
'  - LoadToolingServices for INI-based tool discovery.
'  - GetAvailableTools for current tool registry snapshot.
'  - SharedLibrary.Agents for agent/skill/workspace tool classification.
'  - MultiModelSelectorForm for UI interaction.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods


Partial Public Class ThisAddIn



    ''' <summary>
    ''' Loads tooling service configurations from an INI file and returns tool-capable <see cref="ModelConfig"/> entries.
    ''' </summary>
    ''' <param name="iniPath">INI path containing tool model sections.</param>
    ''' <param name="toolsOnly">When True, filters to entries that have tool-specific prompt/definition fields.</param>
    ''' <returns>List of available tool configurations.</returns>
    Public Function LoadToolingServices(iniPath As String, Optional toolsOnly As Boolean = True) As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()

        ' Register host-internal tool names in the executor registry (idempotent).
        Agents.HostToolRegistration.RegisterWordInternals()

        iniPath = ExpandEnvironmentVariables(iniPath)

        If String.IsNullOrWhiteSpace(iniPath) OrElse Not File.Exists(iniPath) Then
            Return tools
        End If

        Try
            Dim allModels = LoadAlternativeModels(iniPath, _context, StartWithUpcase(ToolFriendlyName), includeToolOnly:=True, toolsOnly:=toolsOnly)

            For Each mc In allModels
                If mc.Deprecated Then Continue For

                If toolsOnly Then
                    If String.IsNullOrWhiteSpace(mc.ToolInstructionsPrompt) AndAlso
                String.IsNullOrWhiteSpace(mc.ToolDefinition) Then
                        Continue For
                    End If
                End If

                mc.Tool = True
                tools.Add(mc)

                ' Register transport-backed external tools that have an APICall template.
                Dim apiTemplate As String =
             If(Not String.IsNullOrWhiteSpace(mc.ToolAPICall), mc.ToolAPICall, mc.APICall)
                If Not String.IsNullOrWhiteSpace(apiTemplate) AndAlso
            Not String.IsNullOrWhiteSpace(mc.ToolName) Then
                    Agents.ToolExecutorRegistry.RegisterExternal(
                 Agents.ToolingHostKind.Word, mc.ToolName)
                End If
            Next

        Catch ex As Exception
            Debug.WriteLine($"LoadToolingServices error: {ex.Message}")
            ToolingFileLogger.LogError("LoadToolingServices error.", ex:=ex)
        End Try

        Return tools
    End Function

    ''' <summary>
    ''' Shows the tool selection dialog and persists the selected tool names into <c>My.Settings.SelectedToolNames</c>.
    ''' </summary>
    ''' <param name="availableTools">List of available tool configurations.</param>
    ''' <param name="preselectAll">Unused parameter in this method body (caller passes a value).</param>
    ''' <returns>Selected tools when the dialog result is OK; otherwise Nothing.</returns>
    Public Function ShowToolSelectionDialog(availableTools As List(Of ModelConfig), Optional preselectAll As Boolean = True, Optional FriendlyName As String = "Tools") As List(Of ModelConfig)
        Dim selectedMainToolNames = GetDiscussInkyEffectiveMainToolNames()
        Dim selectedAdvancedToolNames = GetDiscussInkyEffectiveAdvancedToolNames()
        Dim updatedAdvancedToolNames As List(Of String) = Nothing

        Dim updatedMainToolNames = ShowDiscussInkyToolSelectionDialog(
            selectedMainToolNames,
            selectedAdvancedToolNames,
            updatedAdvancedToolNames)

        If updatedMainToolNames Is Nothing Then
            Return Nothing
        End If

        PersistDiscussInkyToolSelection(
            updatedMainToolNames,
            If(updatedAdvancedToolNames, selectedAdvancedToolNames),
            GetDiscussInkyAdvancedToolsEnabled())

        Dim selected = GetDiscussInkyEffectiveTools(includeImplicitWorkspaceTools:=True)
        SelectedToolNames = selected.Select(Function(t) t.ToolName).Distinct(StringComparer.OrdinalIgnoreCase).ToList()
        Return selected
    End Function

    ''' <summary>
    ''' Returns all available tools by loading external tools from <c>INI_SpecialServicePath</c>,
    ''' adding the internal web tool, conditionally adding the internal search tool
    ''' (only when <c>INI_ISearch</c> is enabled and <c>INI_ISearch_URL</c> is configured),
    ''' and conditionally adding the internal knowledge store search tool
    ''' (only when a knowledge store path is configured and at least one store is indexed).
    ''' </summary>
    ''' <returns>List of available tools.</returns>
    Public Function GetAvailableTools() As List(Of ModelConfig)
        Dim tools As New List(Of ModelConfig)()
        Dim specialServicePath As String = ExpandEnvironmentVariables(INI_SpecialServicePath)

        If Not String.IsNullOrWhiteSpace(specialServicePath) Then
            Dim externalTools = LoadToolingServices(specialServicePath, True)
            tools.AddRange(externalTools)
        End If

        tools.Add(GetInternalWebTool())
        tools.Add(GetInternalDownloadWebFilesTool())

        Dim webGroundingTool =
    SharedLibrary.Agents.WebGroundingTool.Build(
        _context,
        enforcePrivacy:=INI_EnablePrivacyForSearch,
        toolPriority:=997,
        displaySuffix:=InternalToolSuffix)

        If webGroundingTool IsNot Nothing Then
            tools.Add(webGroundingTool)
        End If

        If INI_ISearch AndAlso Not String.IsNullOrWhiteSpace(INI_ISearch_URL) Then
            tools.Add(GetInternalSearchTool(enforcePrivacy:=INI_EnablePrivacyForSearch))
        End If

        tools.AddRange(GetInternalKnowledgeTools())

        ' python_execute: secure sandboxed Python execution.
        ' Only advertised when INI_PythonAgentPath is set, the exe is available, and
        ' (when requested via the packed path) its authenticity has been verified.
        Dim pythonExecuteTool As ModelConfig = Nothing
        If TryConfigureAndBuildPythonExecuteTool(pythonExecuteTool) Then
            tools.Add(pythonExecuteTool)
        End If

        tools.AddRange(SharedLibrary.SharedLibrary.M365ToolService.GetTools(_context, InternalToolSuffix))

        ' Agent layer: session memory, skill loader, and discovered skills/agents (lazy registry-backed).
        Try
            SharedLibrary.Agents.AgentResources.Refresh()
            tools.AddRange(SharedLibrary.Agents.MemoryTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.TextTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.WorkspaceTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.WordTools.BuildAll())
            tools.AddRange(SharedLibrary.Agents.BrowserTools.BuildAll(_context))
            tools.AddRange(SharedLibrary.Agents.WordDocTools.BuildAll())

            Dim jsRunTool As ModelConfig = SharedLibrary.Agents.JsRunTool.Build(_context)
            If jsRunTool IsNot Nothing Then
                tools.Add(jsRunTool)
            End If

            tools.Add(SharedLibrary.Agents.SkillInvokeTool.Build())
            tools.Add(SharedLibrary.Agents.ToolDescribeTool.Build())
            tools.Add(SharedLibrary.Agents.ContextExpandTool.Build())
            tools.Add(SharedLibrary.Agents.ContextCompactTool.Build())
            tools.Add(SharedLibrary.Agents.AskUserTool.Build())

            Dim __agentReg As New SharedLibrary.Agents.ToolRegistry()
            SharedLibrary.Agents.ToolRegistryBuilder.AddSkills(__agentReg, SharedLibrary.Agents.AgentResources.Skills)
            SharedLibrary.Agents.ToolRegistryBuilder.AddAgents(__agentReg, SharedLibrary.Agents.AgentResources.Agents)
            tools.AddRange(__agentReg.MaterializeAll())
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Agent layer registration failed.", ex:=ex)
        End Try

        Return tools
    End Function

    ''' <summary>
    ''' Loads persisted tool selection from <c>My.Settings.SelectedToolNames</c> into <c>SelectedToolNames</c>.
    ''' </summary>
    Public Sub LoadPersistedToolSelection()
        Try
            SelectedToolNames = GetDiscussInkyEffectiveTools(includeImplicitWorkspaceTools:=True).
                Select(Function(t) t.ToolName).
                Distinct(StringComparer.OrdinalIgnoreCase).
                ToList()
        Catch ex As Exception
            SelectedToolNames = New List(Of String)()
            ToolingFileLogger.LogWarn("Failed to load persisted tool selection.", ex:=ex)
        End Try
    End Sub

    ''' <summary>
    ''' Selects tools for the current session either by reusing persisted selections or by showing the tool selection dialog.
    ''' </summary>
    ''' <param name="forceDialog">If True, always shows the selection dialog.</param>
    ''' <returns>Selected tool configurations, or Nothing when the dialog is canceled or no tools are available.</returns>
    Public Function SelectToolsForSession(Optional forceDialog As Boolean = False, Optional FriendlyName As String = ToolFriendlyName) As List(Of ModelConfig)
        Dim selected = GetDiscussInkyEffectiveTools(includeImplicitWorkspaceTools:=True)

        If Not forceDialog AndAlso (selected.Count > 0 OrElse IsDiscussInkyWorkspaceConnected()) Then
            Return selected
        End If

        Return ShowToolSelectionDialog(GetAvailableTools(), preselectAll:=selected.Count = 0, FriendlyName:=FriendlyName)
    End Function

    Private Const AdvancedToolsEnabledSettingName As String = "AdvancedToolsEnabled"

    ' Opt-out persistence: we store only the tools the user has *explicitly deselected*.
    ' Anything not listed is on by default (including newly added tools), so the agentic
    ' platform works with as few clicks as possible while still allowing the user to
    ' deselect some or all tools. New keys (vs. the legacy "Selected..." keys) ensure old
    ' opt-in selections are ignored rather than mis-read as deselections.
    Private Const DeselectedMainToolNamesSettingName As String = "DeselectedMainToolNames"
    Private Const DeselectedAdvancedToolNamesSettingName As String = "DeselectedAdvancedToolNames"





    Private Function IsDiscussInkyWorkspaceConnected() As Boolean
        Try
            Dim ws = SharedLibrary.Agents.WorkspaceStore.Load("word")
            Return ws IsNot Nothing AndAlso
                   Not String.IsNullOrWhiteSpace(ws.RootPath) AndAlso
                   Directory.Exists(ws.RootPath)
        Catch
            Return False
        End Try
    End Function

    Private Function NormalizeDiscussInkyAdvancedToolNames(selectedAdvancedToolNames As IEnumerable(Of String)) As List(Of String)
        Dim result As New List(Of String)(
            If(selectedAdvancedToolNames, Enumerable.Empty(Of String)()).
                Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
                Select(Function(n) n.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase))

        result = result.
            Where(Function(name) Not SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(name)).
            ToList()

        If IsDiscussInkyWorkspaceConnected() Then
            result.AddRange(
                SharedLibrary.Agents.WorkspaceTools.BuildAll().
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    Select(Function(t) t.ToolName))
        End If

        Return result.
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function IsDiscussInkyAdvancedToolName(toolName As String) As Boolean
        If String.IsNullOrWhiteSpace(toolName) Then Return False

        If toolName.StartsWith("skill_", StringComparison.OrdinalIgnoreCase) OrElse
           toolName.StartsWith("agent_", StringComparison.OrdinalIgnoreCase) Then
            Return False
        End If

        If toolName.Equals(InternalWebToolName, StringComparison.OrdinalIgnoreCase) OrElse
           toolName.Equals(InternalSearchToolName, StringComparison.OrdinalIgnoreCase) OrElse
           IsInternalKnowledgeToolName(toolName) OrElse
           SharedLibrary.SharedLibrary.M365ToolService.IsM365ToolName(toolName) Then
            Return False
        End If

        If SharedLibrary.Agents.MemoryTools.IsMemoryTool(toolName) OrElse
           SharedLibrary.Agents.TextTools.IsTextTool(toolName) OrElse
           SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(toolName) OrElse
           SharedLibrary.Agents.WordTools.IsWordTool(toolName) OrElse
           SharedLibrary.Agents.WordDocTools.IsWordDocTool(toolName) OrElse
           SharedLibrary.Agents.BrowserTools.IsBrowserTool(toolName) OrElse
           SharedLibrary.Agents.JsRunTool.IsJsTool(toolName) OrElse
           SharedLibrary.Agents.PythonExecuteTool.IsPythonTool(toolName) OrElse
           SharedLibrary.Agents.ToolDescribeTool.IsDescribeTool(toolName) OrElse
           toolName.Equals(SharedLibrary.Agents.SkillInvokeTool.ToolName, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        Return False
    End Function



    Public Function GetDiscussInkyMainSelectableTools() As List(Of ModelConfig)
        Return DeduplicateToolsByName(
            GetAvailableTools().
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  Not IsDiscussInkyAdvancedToolName(t.ToolName) AndAlso
                                  Not SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(t.ToolName)))
    End Function

    Public Function GetDiscussInkyAdvancedSelectableTools() As List(Of ModelConfig)
        Return DeduplicateToolsByName(
            GetAvailableTools().
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  IsDiscussInkyAdvancedToolName(t.ToolName)))
    End Function

    Public Function GetDiscussInkyAdvancedToolsEnabled() As Boolean
        Return GetWordSettingBoolean(AdvancedToolsEnabledSettingName, False)
    End Function

    Private Function GetDiscussInkyDeselectedMainToolNames() As List(Of String)
        Return SplitPersistedToolNames(GetWordSettingString(DeselectedMainToolNamesSettingName))
    End Function

    Private Function GetDiscussInkyDeselectedAdvancedToolNames() As List(Of String)
        Return SplitPersistedToolNames(GetWordSettingString(DeselectedAdvancedToolNamesSettingName))
    End Function

    ''' <summary>
    ''' Advanced tools available for the current workspace state. Workspace tools are only
    ''' available (and therefore default-on) while a workspace is connected.
    ''' </summary>
    Private Function GetDiscussInkyAvailableAdvancedToolsForState() As List(Of ModelConfig)
        Dim workspaceConnected As Boolean = IsDiscussInkyWorkspaceConnected()

        Return DeduplicateToolsByName(
            GetDiscussInkyAdvancedSelectableTools().
                Where(Function(t) t IsNot Nothing AndAlso
                                  Not String.IsNullOrWhiteSpace(t.ToolName) AndAlso
                                  (workspaceConnected OrElse
                                   Not SharedLibrary.Agents.WorkspaceTools.IsWorkspaceTool(t.ToolName))))
    End Function

    ''' <summary>
    ''' Effective main tool names: every available main tool except those the user deselected.
    ''' </summary>
    Public Function GetDiscussInkyEffectiveMainToolNames() As List(Of String)
        Dim deselected = BuildToolNameSet(GetDiscussInkyDeselectedMainToolNames())

        Return GetDiscussInkyMainSelectableTools().
            Where(Function(t) Not deselected.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    ''' <summary>
    ''' Effective advanced tool names: every state-available advanced tool except those deselected.
    ''' </summary>
    Public Function GetDiscussInkyEffectiveAdvancedToolNames() As List(Of String)
        Dim deselected = BuildToolNameSet(GetDiscussInkyDeselectedAdvancedToolNames())

        Return GetDiscussInkyAvailableAdvancedToolsForState().
            Where(Function(t) Not deselected.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function

    Private Function GetDiscussInkySelectedSkillAllowedToolNames(selectedToolNames As IEnumerable(Of String)) As List(Of String)
        Dim result As New List(Of String)()

        Try
            Dim selectedSet As New HashSet(Of String)(
                If(selectedToolNames, Enumerable.Empty(Of String)()).
                    Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
                    Select(Function(n) n.Trim()),
                StringComparer.OrdinalIgnoreCase)

            If selectedSet.Count = 0 Then
                Return result
            End If

            SharedLibrary.Agents.AgentResources.Refresh()

            For Each skill As SharedLibrary.Agents.SkillDescriptor In SharedLibrary.Agents.AgentResources.Skills
                If skill Is Nothing OrElse String.IsNullOrWhiteSpace(skill.Name) Then
                    Continue For
                End If

                Dim skillToolName As String = "skill_" & skill.Name.Trim()

                If Not selectedSet.Contains(skillToolName) Then
                    Continue For
                End If

                If skill.AllowedTools Is Nothing Then
                    Continue For
                End If

                For Each rawName As String In skill.AllowedTools
                    Dim toolName As String = If(rawName, "").Trim()
                    If toolName <> "" Then
                        result.Add(toolName)
                    End If
                Next
            Next
        Catch ex As Exception
            ToolingFileLogger.LogWarn("Failed to expand selected skill allowed-tools for Discuss Inky.", ex:=ex)
        End Try

        Return result.
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()
    End Function
    Public Function GetDiscussInkyEffectiveTools(Optional includeImplicitWorkspaceTools As Boolean = True) As List(Of ModelConfig)
        Dim advancedEnabled = GetDiscussInkyAdvancedToolsEnabled()

        ' Opt-out defaults: every available tool is on unless the user deselected it.
        Dim mainNames = GetDiscussInkyEffectiveMainToolNames()
        Dim advancedNames As List(Of String) =
            If(advancedEnabled, GetDiscussInkyEffectiveAdvancedToolNames(), New List(Of String)())

        Dim result As New List(Of ModelConfig)()
        Dim mainSet = BuildToolNameSet(mainNames)
        Dim advancedSet = BuildToolNameSet(advancedNames)

        For Each tool In GetDiscussInkyMainSelectableTools()
            If mainSet.Contains(tool.ToolName) Then
                result.Add(tool)
            End If
        Next

        If advancedEnabled Then
            For Each tool In GetDiscussInkyAvailableAdvancedToolsForState()
                If advancedSet.Contains(tool.ToolName) Then
                    result.Add(tool)
                End If
            Next
        End If

        Dim explicitlySelectedToolNames As New List(Of String)()
        explicitlySelectedToolNames.AddRange(mainNames)
        explicitlySelectedToolNames.AddRange(advancedNames)

        Dim skillRequiredToolNames As List(Of String) =
        GetDiscussInkySelectedSkillAllowedToolNames(explicitlySelectedToolNames)

        If skillRequiredToolNames.Count > 0 Then
            Dim requiredSet = BuildToolNameSet(skillRequiredToolNames)

            For Each tool In GetAvailableTools()
                If tool Is Nothing OrElse String.IsNullOrWhiteSpace(tool.ToolName) Then Continue For

                If requiredSet.Contains(tool.ToolName) Then
                    result.Add(tool)
                End If
            Next
        End If

        result = DeduplicateToolsByName(result)

        SelectedToolNames = result.Select(Function(t) t.ToolName).ToList()
        Return result
    End Function

    Public Sub PersistDiscussInkyToolSelection(selectedMainToolNames As IEnumerable(Of String),
                                               selectedAdvancedToolNames As IEnumerable(Of String),
                                               advancedToolsEnabled As Boolean)
        ' Persist the *deselected* tools (opt-out): available-minus-checked, per category.
        Dim mainChecked = BuildToolNameSet(selectedMainToolNames)
        Dim deselectedMain = GetDiscussInkyMainSelectableTools().
            Where(Function(t) Not mainChecked.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        Dim advancedChecked = BuildToolNameSet(selectedAdvancedToolNames)
        Dim deselectedAdvanced = GetDiscussInkyAvailableAdvancedToolsForState().
            Where(Function(t) Not advancedChecked.Contains(t.ToolName)).
            Select(Function(t) t.ToolName).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        SetWordSettingValue(DeselectedMainToolNamesSettingName, JoinPersistedToolNames(deselectedMain))
        SetWordSettingValue(DeselectedAdvancedToolNamesSettingName, JoinPersistedToolNames(deselectedAdvanced))
        SetWordSettingValue(AdvancedToolsEnabledSettingName, advancedToolsEnabled)

        Dim effective = GetDiscussInkyEffectiveTools(includeImplicitWorkspaceTools:=False)
        My.Settings.SelectedToolNames = String.Join("|", effective.Select(Function(t) t.ToolName).Distinct(StringComparer.OrdinalIgnoreCase))
        My.Settings.Save()
    End Sub

    Private Function ShowDiscussInkyAdvancedToolSelectionDialog(selectedAdvancedToolNames As IEnumerable(Of String)) As List(Of String)
        Dim availableTools = GetDiscussInkyAvailableAdvancedToolsForState()
        Dim preselected = If(selectedAdvancedToolNames, Enumerable.Empty(Of String)()).
            Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
            Select(Function(n) n.Trim()).
            Distinct(StringComparer.OrdinalIgnoreCase).
            ToList()

        Using selector As New MultiModelSelectorForm(
            availableTools,
            "",
            $"{AN} - Select Advanced Tools",
            resetChecked:=False,
            preselectMany:=preselected,
            instruction:="Select the advanced tools that may be callable. All available tools are on by default; " &
                         "uncheck any you want to disable. Workspace tools appear here only while a workspace is connected.")

            If selector.ShowDialog() = DialogResult.OK Then
                Return selector.SelectedModels.
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    Select(Function(t) t.ToolName).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()
            End If
        End Using

        Return Nothing
    End Function

    Public Function ShowDiscussInkyToolSelectionDialog(selectedMainToolNames As IEnumerable(Of String),
                                                       selectedAdvancedToolNames As IEnumerable(Of String),
                                                       ByRef updatedAdvancedToolNames As List(Of String)) As List(Of String)

        Dim availableTools = GetDiscussInkyMainSelectableTools()
        Dim workingAdvanced As New List(Of String)(
            If(selectedAdvancedToolNames, Enumerable.Empty(Of String)()).
                Where(Function(n) Not String.IsNullOrWhiteSpace(n)).
                Select(Function(n) n.Trim()).
                Distinct(StringComparer.OrdinalIgnoreCase))

        Using selector As New MultiModelSelectorForm(
            availableTools,
            "",
            $"{AN} - Select {ToolFriendlyName}",
            resetChecked:=False,
            preselectMany:=If(selectedMainToolNames, New List(Of String)()),
            instruction:="Select the agents, sources, skills, and connector-oriented tools you want to make available to the model. " &
                         "Note: all available tools are on by default (including newly added ones); uncheck any you want to disable. " &
                         "Advanced tools are managed separately through the 'Advanced tools…' button.")

            selector.AddExtraButton("Advanced tools…",
                Sub(s, e)
                    Dim advanced = ShowDiscussInkyAdvancedToolSelectionDialog(workingAdvanced)
                    If advanced IsNot Nothing Then
                        workingAdvanced = advanced
                    End If
                End Sub)

            selector.AddExtraButton("Skills && Agents…",
                Sub(s, e)
                    Using f As New SharedLibrary.Agents.AgentResourcesViewerForm(_context)
                        f.ShowDialog(selector)
                    End Using
                End Sub)

            selector.AddExtraButton("Memory…",
                Sub(s, e)
                    Using f As New SharedLibrary.Agents.SessionMemoryViewerForm()
                        f.ShowDialog(selector)
                    End Using
                End Sub)

            selector.AddExtraButton("Workspace…",
                Sub(s, e)
                    Using f As New WordWorkspaceForm()
                        f.ShowDialog(selector)
                    End Using
                End Sub)

            If selector.ShowDialog() = DialogResult.OK Then
                updatedAdvancedToolNames = workingAdvanced.
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

                Dim selectedMain = selector.SelectedModels.
                    Where(Function(t) t IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(t.ToolName)).
                    Select(Function(t) t.ToolName).
                    Distinct(StringComparer.OrdinalIgnoreCase).
                    ToList()

                Return selectedMain
            End If
        End Using

        Return Nothing
    End Function

    Public Function SelectDiscussInkyToolsForSession(Optional forceDialog As Boolean = False) As List(Of ModelConfig)
        Dim selectedMain = GetDiscussInkyEffectiveMainToolNames()
        Dim selectedAdvanced = GetDiscussInkyEffectiveAdvancedToolNames()

        If Not forceDialog Then
            Dim effective = GetDiscussInkyEffectiveTools()
            If effective.Count > 0 OrElse IsDiscussInkyWorkspaceConnected() Then
                Return effective
            End If
        End If

        Dim updatedAdvanced As List(Of String) = Nothing
        Dim updatedMain = ShowDiscussInkyToolSelectionDialog(selectedMain, selectedAdvanced, updatedAdvanced)

        If updatedMain Is Nothing Then
            Return Nothing
        End If

        PersistDiscussInkyToolSelection(updatedMain, If(updatedAdvanced, selectedAdvanced), GetDiscussInkyAdvancedToolsEnabled())
        Return GetDiscussInkyEffectiveTools()
    End Function



End Class
