' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.MCPImporter.vb
' Purpose: Imports tool metadata from MCP servers and generates Special Service
'          INI sections for selected tools, including transport probing,
'          schema mapping, response-path inference, and user-driven export.
'
' Architecture:
'  - Endpoint discovery:
'      - Probes Streamable HTTP endpoints first.
'      - Falls back to legacy SSE transport when needed.
'  - MCP protocol flow:
'      - Sends JSON-RPC `initialize` and `tools/list`.
'      - Supports paged tool discovery via cursor-based continuation.
'  - Tool inspection and mapping:
'      - Executes sample calls to infer response paths.
'      - Maps JSON Schema parameters to INI `Parameter1..4` fields and preserves
'        overflow metadata in generated tool definitions.
'  - User workflow:
'      - Collects auth/header settings and optional live probe credentials.
'      - Lets the user choose tools, output mode, timeout, and output file.
'  - INI generation:
'      - Emits `APICall`, `ToolAPICall`, defaults, tool instructions, and tool
'        definitions in INI-safe single-line form.
'
' Dependencies:
'  - System.Net.Http for MCP transport requests.
'  - Newtonsoft.Json for payload construction and parsing.
'  - SharedLibrary UI helpers for dialogs and selection forms.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.IO
Imports System.Net.Http
Imports System.Text
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ' ─────────────────────────────────────────────────────────────────────────
    '  MCP → INI  constants
    ' ─────────────────────────────────────────────────────────────────────────

    Private Const MCP_DEFAULT_TIMEOUT As Integer = 60000
    Private Const MCP_MAX_INI_PARAMS As Integer = 4
    Private Const MCP_PROTOCOL_VERSION As String = "2025-03-26"

    ''' <summary>Common Streamable-HTTP sub-paths to probe when the root URL fails.</summary>
    Private Shared ReadOnly MCP_PATH_SUFFIXES As String() = {"/mcp", "/rpc", "/json-rpc", "/api/mcp"}

    ''' <summary>Common SSE-transport sub-paths to probe.</summary>
    Private Shared ReadOnly MCP_SSE_SUFFIXES As String() = {"/sse", "/events"}

    ' ─────────────────────────────────────────────────────────────────────────
    '  MCP transport session state (SSE transport only)
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Holds state for an active SSE-transport MCP session.
    ''' For SSE transport, a GET /sse connection must be opened first to receive the
    ''' session endpoint URL, then JSON-RPC POSTs go to that endpoint. Responses
    ''' arrive as SSE events on the GET stream.
    ''' </summary>
    Private Class MCPSSESession
        Implements IDisposable

        ''' <summary>The base SSE URL (e.g. https://server/sse).</summary>
        Public Property SseUrl As String = ""

        ''' <summary>The POST endpoint URL returned by the server in the "endpoint" SSE event.</summary>
        Public Property PostEndpoint As String = ""

        ''' <summary>The HttpClient kept alive for the SSE stream.</summary>
        Public Property Client As HttpClient = Nothing

        ''' <summary>The SSE response stream reader.</summary>
        Public Property StreamReader As StreamReader = Nothing

        ''' <summary>The underlying HTTP response to dispose.</summary>
        Public Property SseResponse As HttpResponseMessage = Nothing

        Public Property IsConnected As Boolean = False

        Public Sub Dispose() Implements IDisposable.Dispose
            Try : StreamReader?.Dispose() : Catch : End Try
            Try : SseResponse?.Dispose() : Catch : End Try
            Try : Client?.Dispose() : Catch : End Try
            IsConnected = False
        End Sub
    End Class

    ''' <summary>Active SSE session, Nothing when using Streamable HTTP transport.</summary>
    Private _mcpSseSession As MCPSSESession = Nothing

    ' ─────────────────────────────────────────────────────────────────────────
    '  Entry point – called from ribbon / menu
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Interactive MCP import workflow:
    ''' discovers server tools and writes generated INI sections to a user-selected file.
    ''' </summary>
    Public Async Sub ImportMCPServer()
        Try
            ' ── 1. Collect MCP endpoint URL ──────────────────────────────────
            Dim mcpUrl As String = ShowCustomInputBox(
                "Enter the MCP server URL (Streamable HTTP or SSE endpoint):",
                $"{AN} MCP Import", True).Trim()

            If String.IsNullOrWhiteSpace(mcpUrl) Then Return
            mcpUrl = mcpUrl.TrimEnd("/"c)

            ' ── 2. Initialize + discover tools; if needed, bootstrap OAuth and retry ──
            Dim authParams As New MCPAuthInfo()

            Dim tools As List(Of MCPToolInfo) = Nothing
            Dim serverName As String = ""
            Dim resolvedUrl As String = mcpUrl
            Dim sseBaseUrl As String = ""
            Dim detectedTransport As String = ""

            Dim initResult As MCPInitResult = Nothing
            Dim firstDiscoveryError As Exception = Nothing

            Try
                initResult = Await MCPInitializeWithDiscovery(mcpUrl, authParams)
                resolvedUrl = initResult.ResolvedUrl
                sseBaseUrl = initResult.SseBaseUrl
                serverName = initResult.ServerName
                detectedTransport = initResult.Transport
                tools = Await MCPListTools(resolvedUrl, authParams)
            Catch ex As Exception
                firstDiscoveryError = ex
            Finally
                If _mcpSseSession IsNot Nothing Then
                    _mcpSseSession.Dispose()
                    _mcpSseSession = Nothing
                End If
            End Try

            If firstDiscoveryError IsNot Nothing Then
                Dim discoveredOAuth As SLib.MCPProtectedResourceOAuthResult = Nothing
                Dim oauthBootstrapError As Exception = Nothing

                SLib.SetMCPOAuthDebugEnabled(INI_APIDebug)

                Try
                    discoveredOAuth = Await SLib.AcquireMCPProtectedResourceOAuthAsync(mcpUrl).ConfigureAwait(False)
                Catch ex As Exception
                    oauthBootstrapError = ex
                End Try

                If oauthBootstrapError IsNot Nothing Then
                    Await SwitchToUi(
                        Sub()
                            ShowCustomMessageBox(
                                "Failed to query the MCP server before OAuth bootstrap:" & vbCrLf & vbCrLf &
                                firstDiscoveryError.Message & vbCrLf & vbCrLf &
                                "Automatic OAuth bootstrap also failed:" & vbCrLf & vbCrLf &
                                oauthBootstrapError.Message,
                                $"{AN} MCP Import")
                        End Sub)
                    Return
                End If

                If discoveredOAuth Is Nothing Then
                    Await SwitchToUi(
                        Sub()
                            ShowCustomMessageBox(
                                "Failed to query the MCP server:" & vbCrLf & vbCrLf &
                                firstDiscoveryError.Message & vbCrLf & vbCrLf &
                                "No OAuth challenge or protected-resource metadata could be discovered at the MCP URL.",
                                $"{AN} MCP Import")
                        End Sub)
                    Return
                End If

                ApplyDiscoveredMCPOAuth(authParams, discoveredOAuth)

                Await SwitchToUi(
                    Sub()
                        ShowCustomMessageBox(
                            "The MCP server requires OAuth authentication." & vbCrLf & vbCrLf &
                            "Authentication completed successfully. The importer will now retry server discovery.",
                            $"{AN} MCP Import")
                    End Sub)

                Dim secondDiscoveryError As Exception = Nothing
                initResult = Nothing
                tools = Nothing
                serverName = ""
                resolvedUrl = mcpUrl
                sseBaseUrl = ""
                detectedTransport = ""

                Try
                    initResult = Await MCPInitializeWithDiscovery(mcpUrl, authParams)
                    resolvedUrl = initResult.ResolvedUrl
                    sseBaseUrl = initResult.SseBaseUrl
                    serverName = initResult.ServerName
                    detectedTransport = initResult.Transport
                    tools = Await MCPListTools(resolvedUrl, authParams)
                Catch ex As Exception
                    secondDiscoveryError = ex
                Finally
                    If _mcpSseSession IsNot Nothing Then
                        _mcpSseSession.Dispose()
                        _mcpSseSession = Nothing
                    End If
                End Try

                If secondDiscoveryError IsNot Nothing Then
                    Await SwitchToUi(
                        Sub()
                            ShowCustomMessageBox(
                                "OAuth authentication succeeded, but querying the MCP server still failed:" & vbCrLf & vbCrLf &
                                secondDiscoveryError.Message,
                                $"{AN} MCP Import")
                        End Sub)
                    Return
                End If
            End If

            If tools Is Nothing OrElse tools.Count = 0 Then
                Await SwitchToUi(Sub() ShowCustomMessageBox("The MCP server did not return any tools."))
                Return
            End If

            ' ── All UI from here must be on the STA thread ───────────────────
            Dim selectedTools As List(Of MCPToolInfo) = Nothing
            Dim responseKey As String = ""
            Dim mergePrompt As String = ""
            Dim timeoutVal As Long = 200000
            Dim normalizeSchemas As Boolean = True
            Dim cancelled As Boolean = False

            Await SwitchToUi(Sub()
                                 If Not resolvedUrl.Equals(mcpUrl, StringComparison.OrdinalIgnoreCase) Then
                                     ShowCustomMessageBox(
                                         $"The MCP endpoint was resolved to:{vbCrLf}{resolvedUrl}{vbCrLf}" &
                                         $"Transport: {detectedTransport}{vbCrLf}{vbCrLf}" &
                                         $"This URL will be used in the generated INI file.")
                                 End If

                                 ' ── 4. Let user select which tools to import ─────────
                                 selectedTools = ShowMCPToolSelectionDialog(tools, serverName)
                                 If selectedTools Is Nothing OrElse selectedTools.Count = 0 Then
                                     cancelled = True
                                     Return
                                 End If

                                 Dim normalizeChoice = ShowCustomYesNoBox(
                                     "Should the imported ToolDefinition schemas be normalized for broad model compatibility?" & vbCrLf & vbCrLf &
                                     "• Normalize: flattens or removes advanced JSON Schema features such as $ref, $defs, anyOf/oneOf, titles, defaults, and nullable unions. " &
                                     "Recommended for Gemini and simpler tool-calling models." & vbCrLf &
                                     "• Keep original: preserves the MCP schema exactly as returned by the server. " &
                                     "Use this when targeting models that support richer JSON Schema." & vbCrLf & vbCrLf &
                                     "You can re-import later using the other mode if needed.",
                                     "Normalize",
                                     "Keep original",
                                     $"{AN} MCP Import")

                                 normalizeSchemas = (normalizeChoice <> 2)
                             End Sub)

            If cancelled OrElse selectedTools Is Nothing OrElse selectedTools.Count = 0 Then Return

            ' ── 5. Detect response key ───────────────────────────────────────
            Dim detectedResponseKey As String = "result.content[0].text"
            If detectedTransport = "streamable-http" Then
                detectedResponseKey = Await ProbeMCPResponseKey(resolvedUrl, authParams, selectedTools(0))
                If String.IsNullOrWhiteSpace(detectedResponseKey) Then
                    detectedResponseKey = "result.content[0].text"
                End If
            End If

            Await SwitchToUi(Sub()
                                 responseKey = ShowCustomInputBox(
                                     "The following JSON path will be used for extracting tool responses." & vbCrLf &
                                     "Adjust if needed (this becomes the INI 'Response' key):",
                                     $"{AN} MCP Import", True, detectedResponseKey).Trim()

                                 If String.IsNullOrWhiteSpace(responseKey) Then responseKey = detectedResponseKey

                                 ' ── 6. Collect per-server optional settings ──────────
                                 Dim optParams As New List(Of InputParameter)()
                                 optParams.Add(New InputParameter("MergePrompt (optional, leave blank to skip)", ""))
                                 optParams.Add(New InputParameter("Timeout (ms)", 200000))

                                 Dim optArray() As InputParameter = optParams.ToArray()
                                 If Not ShowCustomVariableInputForm("Configure additional settings:", $"{AN} MCP Import", optArray) Then
                                     cancelled = True
                                     Return
                                 End If

                                 mergePrompt = If(optArray(0).Value?.ToString(), "").Trim()
                                 If TypeOf optArray(1).Value Is Integer Then
                                     timeoutVal = CLng(CInt(optArray(1).Value))
                                 Else
                                     Long.TryParse(If(optArray(1).Value?.ToString(), "200000"), timeoutVal)
                                 End If
                             End Sub)

            If cancelled Then Return

            ' ── 7. Build INI sections (no UI needed) ─────────────────────────
            Dim sb As New StringBuilder()
            sb.AppendLine("; MCP Server: " & resolvedUrl)
            sb.AppendLine($"; Server: {serverName}")
            sb.AppendLine($"; Transport: {detectedTransport}")
            sb.AppendLine($"; Protocol: {MCP_PROTOCOL_VERSION}")
            sb.AppendLine($"; Schema mode: {If(normalizeSchemas, "normalized", "original")}")
            sb.AppendLine($"; Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            sb.AppendLine($"; Tools imported: {selectedTools.Count} of {tools.Count} discovered (using Freestyle 'MCP' command)")
            If detectedTransport = "sse" Then
                sb.AppendLine(";")
                sb.AppendLine("; NOTE: This server uses the legacy SSE transport. The Endpoint uses the")
                sb.AppendLine($"; '{MCP_SSE_PREFIX}' prefix to signal that a session handshake is required.")
            End If
            sb.AppendLine()

            Dim paramWarnings As New List(Of String)()
            Dim toolOnly As Boolean = True

            ' Ask the user whether tools should be Tool Only or also usable as Special Services
            Await SwitchToUi(Sub()
                                 Dim choice = ShowCustomYesNoBox(
                                     "Should the imported tools be marked as 'Tool Only'?" & vbCrLf & vbCrLf &
                                     "• Tool Only (ToolOnly=True): Tools can only be invoked via the tooling pipeline." & vbCrLf &
                                     "• Tool + Special Service (ToolOnly=False): Tools are also available as standalone Special Services.",
                                     "Tool Only", "Tool + Special Service", $"{AN} MCP Import")
                                 toolOnly = (choice <> 2)
                             End Sub)

            For Each tool In selectedTools
                Dim prefix As String = If(Not String.IsNullOrWhiteSpace(serverName), serverName & " ", "")
                Dim sectionName As String = prefix & MCPHumanizeName(tool.Name)

                Dim section As String = BuildINISectionForTool(
                    tool, sectionName, resolvedUrl, sseBaseUrl, authParams,
                    responseKey, mergePrompt, timeoutVal, paramWarnings, toolOnly, normalizeSchemas)

                sb.AppendLine(section)
            Next

            ' ── 8. Save file + show warnings (all UI on STA thread) ──────────
            Await SwitchToUi(Sub()
                                 If paramWarnings.Count > 0 Then
                                     Dim warnText As String =
                                         $"{paramWarnings.Count} tool(s) have more than {MCP_MAX_INI_PARAMS} parameters beyond the primary query. " &
                                         $"Only the first {MCP_MAX_INI_PARAMS} are mapped to Parameter1–4; the rest are available via ToolAPICall/ToolDefinition." &
                                         vbCrLf & vbCrLf & String.Join(vbCrLf, paramWarnings)
                                     ShowCustomMessageBox(warnText, $"{AN} MCP Import")
                                 End If

                                 Dim defaultDir As String = ""
                                 Try
                                     defaultDir = Path.GetDirectoryName(If(INI_SpecialServicePath, ""))
                                 Catch
                                 End Try
                                 If String.IsNullOrWhiteSpace(defaultDir) Then
                                     defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
                                 End If

                                 Dim safeName As String = MCPSanitizeFileName(If(serverName, "mcp_server"))
                                 Dim defaultFile As String = Path.Combine(defaultDir, safeName & "_services.ini")

                                 Using dlg As New SaveFileDialog()
                                     dlg.Title = "Save generated INI file"
                                     dlg.Filter = "INI files (*.ini)|*.ini|All files (*.*)|*.*"
                                     dlg.FileName = Path.GetFileName(defaultFile)
                                     dlg.InitialDirectory = defaultDir

                                     Dim __safeDialogOwner385 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                                     If If(__safeDialogOwner385 IsNot Nothing, dlg.ShowDialog(__safeDialogOwner385), dlg.ShowDialog()) = DialogResult.OK Then
                                         File.WriteAllText(dlg.FileName, sb.ToString(), Encoding.UTF8)
                                         ShowCustomMessageBox(
                                             $"INI file saved to:{vbCrLf}{dlg.FileName}{vbCrLf}{vbCrLf}" &
                                             $"{selectedTools.Count} tool section(s) generated.{vbCrLf}{vbCrLf}" &
                                             $"To use these services, add the file content to your Special Services INI file or set SpecialServicePath to this file in your {AN2}.ini.")
                                     End If
                                 End Using
                             End Sub)

        Catch ex As Exception
            ' Store for re-throw after Try block (Await not allowed in Catch)
            Dim errorToShow As Exception = ex
            Try
                mainThreadControl.Invoke(New MethodInvoker(
                    Sub() MessageBox.Show("Error in ImportMCPServer: " & errorToShow.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)))
            Catch
                MessageBox.Show("Error in ImportMCPServer: " & errorToShow.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        Finally
            ' Dispose directly (Await not allowed in Finally)
            If _mcpSseSession IsNot Nothing Then
                _mcpSseSession.Dispose()
                _mcpSseSession = Nothing
            End If
        End Try
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────
    '  Helper data classes
    ' ─────────────────────────────────────────────────────────────────────────

    Private Class MCPAuthInfo
        Public Property APIKeyPlaceholder As String = ""
        Public Property LiveAPIKey As String = ""
        Public Property APIKeyEncrypted As Boolean = False
        Public Property APIKeyPrefix As String = ""
        Public Property HeaderA As String = ""
        Public Property HeaderB As String = ""

        Public Property OAuth2 As Boolean = False
        Public Property OAuth2ClientId As String = ""
        Public Property OAuth2ClientSecret As String = ""
        Public Property OAuth2AuthorizationEndpoint As String = ""
        Public Property OAuth2TokenEndpoint As String = ""
        Public Property OAuth2DeviceAuthorizationEndpoint As String = ""
        Public Property OAuth2Scope As String = ""
        Public Property OAuth2Resource As String = ""
        Public Property OAuth2ExpiresIn As Integer = 3600
        Public Property OAuth2UseDeviceFlow As Boolean = False
    End Class

    Private Class MCPInitResult
        Public Property ServerName As String = ""
        Public Property ProtocolVersion As String = ""
        Public Property ResolvedUrl As String = ""
        ''' <summary>
        ''' For SSE transport, the base SSE URL used for session handshakes (e.g. https://server/sse).
        ''' Empty for Streamable HTTP.
        ''' </summary>
        Public Property SseBaseUrl As String = ""
        ''' <summary>"streamable-http" or "sse"</summary>
        Public Property Transport As String = "streamable-http"
    End Class

    Private Class MCPToolInfo
        Public Property Name As String = ""
        Public Property Description As String = ""
        Public Property InputSchema As JObject = Nothing
    End Class

    Private Class MCPParamInfo
        Public Property Name As String = ""
        Public Property Type As String = "string"
        Public Property Description As String = ""
        Public Property IsRequired As Boolean = False
        Public Property DefaultValue As String = ""
        Public Property EnumValues As List(Of String) = Nothing
        Public Property Minimum As Integer? = Nothing
        Public Property Maximum As Integer? = Nothing
    End Class

    ' ─────────────────────────────────────────────────────────────────────────
    '  Auth collection
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Collects MCP authentication/header settings from the user.
    ''' </summary>
    ''' <returns>
    ''' Auth settings object, or <c>Nothing</c> when user cancels.
    ''' </returns>
    Private Function CollectMCPAuthInfo() As MCPAuthInfo
        Dim paramDefs As New List(Of InputParameter)()
        paramDefs.Add(New InputParameter("API Key placeholder for INI (e.g. [[Your API Key]], or blank for open servers)", ""))
        paramDefs.Add(New InputParameter("Actual API Key for probing the server now (blank for open servers)", ""))
        paramDefs.Add(New InputParameter("API Key Encrypted?", False))
        paramDefs.Add(New InputParameter("API Key Prefix (leave blank if none)", ""))
        paramDefs.Add(New InputParameter("Header Name (e.g. Authorization, or blank for no auth)", ""))
        paramDefs.Add(New InputParameter("Header Value template (use {apikey} as placeholder, or blank)", ""))

        Dim params() As InputParameter = paramDefs.ToArray()
        If Not ShowCustomVariableInputForm("Configure MCP server authentication (leave all blank for open servers):", $"{AN} MCP Import", params) Then
            Return Nothing
        End If

        Dim info As New MCPAuthInfo()
        info.APIKeyPlaceholder = If(params(0).Value?.ToString(), "").Trim()
        info.LiveAPIKey = If(params(1).Value?.ToString(), "").Trim()
        info.APIKeyEncrypted = If(TypeOf params(2).Value Is Boolean, CBool(params(2).Value), False)
        info.APIKeyPrefix = If(params(3).Value?.ToString(), "").Trim()
        info.HeaderA = If(params(4).Value?.ToString(), "").Trim()
        info.HeaderB = If(params(5).Value?.ToString(), "").Trim()

        If String.IsNullOrWhiteSpace(info.LiveAPIKey) AndAlso
           Not String.IsNullOrWhiteSpace(info.APIKeyPlaceholder) AndAlso
           Not info.APIKeyPlaceholder.StartsWith("[[") Then
            info.LiveAPIKey = info.APIKeyPlaceholder
        End If
        Return info
    End Function

    Private Shared Sub ApplyDiscoveredMCPOAuth(auth As MCPAuthInfo, oauth As SLib.MCPProtectedResourceOAuthResult)
        auth.OAuth2 = True
        auth.OAuth2ClientId = oauth.ClientId
        auth.OAuth2ClientSecret = oauth.ClientSecret
        auth.OAuth2AuthorizationEndpoint = oauth.AuthorizationEndpoint
        auth.OAuth2TokenEndpoint = oauth.TokenEndpoint
        auth.OAuth2DeviceAuthorizationEndpoint = oauth.DeviceAuthorizationEndpoint
        auth.OAuth2Scope = oauth.Scope
        auth.OAuth2Resource = oauth.Resource
        auth.OAuth2ExpiresIn = oauth.ExpiresIn
        auth.OAuth2UseDeviceFlow = oauth.UseDeviceFlow

        auth.APIKeyPlaceholder = oauth.ClientSecret
        auth.LiveAPIKey = oauth.AccessToken
        auth.APIKeyEncrypted = False
        auth.APIKeyPrefix = ""
        auth.HeaderA = "Authorization"
        auth.HeaderB = "Bearer {apikey}"

        ' Build the OAuth2Endpoint exactly the same way it will be stored in the INI,
        ' so the runtime cache key matches.
        Dim oauthEndpointValue As String

        If oauth.UseDeviceFlow Then
            oauthEndpointValue = "device:" & oauth.DeviceAuthorizationEndpoint & "¦" & oauth.TokenEndpoint
        Else
            oauthEndpointValue = oauth.AuthorizationEndpoint & "¦" & oauth.TokenEndpoint
        End If

        If Not String.IsNullOrWhiteSpace(oauth.Resource) Then
            oauthEndpointValue &= "¦" & oauth.Resource
        End If

        Dim cacheKey As String = SLib.BuildMCPTokenCacheKey(oauth.ClientId, oauthEndpointValue)
        Dim expiresInSeconds As Integer = If(oauth.ExpiresIn > 0, oauth.ExpiresIn, 3600)
        Dim expiryUtc As DateTime = DateTime.UtcNow.AddSeconds(Math.Max(60, expiresInSeconds - 300))

        SLib.StoreCachedMCPToken(cacheKey, oauth.AccessToken, expiryUtc)
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────
    '  Tool selection dialog
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Shows tool-selection UI and returns selected MCP tools.
    ''' </summary>
    ''' <returns>
    ''' Selected tool list, or <c>Nothing</c> when user cancels.
    ''' </returns>
    Private Function ShowMCPToolSelectionDialog(tools As List(Of MCPToolInfo), serverName As String) As List(Of MCPToolInfo)
        Dim tempConfigs As New List(Of ModelConfig)()
        Dim configToTool As New Dictionary(Of ModelConfig, MCPToolInfo)()

        For Each t In tools
            Dim mc As New ModelConfig()
            mc.Model = t.Name
            mc.ModelDescription = If(Not String.IsNullOrWhiteSpace(serverName), serverName & " ", "") &
                                  MCPHumanizeName(t.Name) &
                                  If(Not String.IsNullOrWhiteSpace(t.Description), " – " & t.Description, "")
            mc.ToolName = t.Name
            mc.ToolDefinition = "placeholder"
            mc.ToolInstructionsPrompt = If(t.Description, "")
            mc.Tool = True
            tempConfigs.Add(mc)
            configToTool(mc) = t
        Next

        Dim preselectAll = tempConfigs.Select(Function(c) c.ModelDescription).ToList()
        Dim selector As New MultiModelSelectorForm(
            tempConfigs, "",
            $"{AN} MCP Import – Select tools to import",
            resetChecked:=False,
            preselectMany:=preselectAll,
            instruction:="Select the tools/services you want to generate INI sections for:")

        Dim __safeDialogOwner583 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
        If If(__safeDialogOwner583 IsNot Nothing, selector.ShowDialog(__safeDialogOwner583), selector.ShowDialog()) <> DialogResult.OK Then Return Nothing

        Dim selected = selector.SelectedModels
        If selected Is Nothing OrElse selected.Count = 0 Then Return Nothing

        Dim result As New List(Of MCPToolInfo)()
        For Each mc In selected
            If configToTool.ContainsKey(mc) Then result.Add(configToTool(mc))
        Next
        Return result
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  MCP transport discovery + initialization
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Discovers the transport type and initialises the MCP session.
    ''' Strategy:
    '''   1. Try SSE transport GET to the exact URL.
    '''   2. Try SSE transport GET to common sub-paths (/sse, /events).
    '''   3. Try Streamable HTTP POST to the exact URL.
    '''   4. Try Streamable HTTP POST to common sub-paths (/mcp, /rpc, …).
    ''' </summary>
    Private Async Function MCPInitializeWithDiscovery(mcpUrl As String, auth As MCPAuthInfo) As Task(Of MCPInitResult)
        Dim baseUrl As String = mcpUrl.TrimEnd("/"c)

        ' ── Phase 1: Try SSE transport on the exact URL ──────────────────
        Try
            Dim r = Await MCPInitializeSSE(baseUrl, auth)
            r.Transport = "sse"
            Return r
        Catch
        End Try

        ' ── Phase 2: Try SSE transport on common sub-paths ───────────────
        For Each suffix In MCP_SSE_SUFFIXES
            Try
                Dim r = Await MCPInitializeSSE(baseUrl & suffix, auth)
                r.Transport = "sse"
                Return r
            Catch
            End Try
        Next

        ' ── Phase 3: Try Streamable HTTP on the exact URL ────────────────
        Try
            Dim r = Await MCPInitializeStreamableHTTP(baseUrl, auth)
            r.ResolvedUrl = baseUrl
            r.Transport = "streamable-http"
            Return r
        Catch
        End Try

        ' ── Phase 4: Try Streamable HTTP on common sub-paths ─────────────
        For Each suffix In MCP_PATH_SUFFIXES
            Try
                Dim candidate = baseUrl & suffix
                Dim r = Await MCPInitializeStreamableHTTP(candidate, auth)
                r.ResolvedUrl = candidate
                r.Transport = "streamable-http"
                Return r
            Catch
            End Try
        Next

        Throw New Exception(
            $"Could not connect to MCP server at {mcpUrl}.{vbCrLf}{vbCrLf}" &
            $"Tried SSE transport on: {baseUrl}, {String.Join(", ", MCP_SSE_SUFFIXES.Select(Function(s) baseUrl & s))}{vbCrLf}" &
            $"Tried Streamable HTTP on: {baseUrl}, {String.Join(", ", MCP_PATH_SUFFIXES.Select(Function(s) baseUrl & s))}{vbCrLf}{vbCrLf}" &
            $"Please verify the correct endpoint URL with the server provider.")
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  Streamable HTTP transport
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Initialises via Streamable HTTP: POST the initialize JSON-RPC to the URL.
    ''' </summary>
    Private Async Function MCPInitializeStreamableHTTP(mcpUrl As String, auth As MCPAuthInfo) As Task(Of MCPInitResult)
        Dim payload As New JObject From {
            {"jsonrpc", "2.0"}, {"id", 1}, {"method", "initialize"},
            {"params", New JObject From {
                {"protocolVersion", MCP_PROTOCOL_VERSION},
                {"capabilities", New JObject()},
                {"clientInfo", New JObject From {{"name", AN}, {"version", "1.0"}}}
            }}
        }

        Dim responseJson = Await SendStreamableHTTPRequest(mcpUrl, payload, auth)

        Dim result As New MCPInitResult() With {.ResolvedUrl = mcpUrl}
        Dim serverInfo = responseJson.SelectToken("result.serverInfo")
        If serverInfo IsNot Nothing Then
            result.ServerName = If(serverInfo("name")?.ToString(), "")
            result.ProtocolVersion = If(responseJson.SelectToken("result.protocolVersion")?.ToString(), "")
        End If

        ' Send initialized notification
        Dim notif As New JObject From {{"jsonrpc", "2.0"}, {"method", "notifications/initialized"}}
        Try
            Await SendStreamableHTTPRequest(mcpUrl, notif, auth, expectResponse:=False)
        Catch
        End Try
        Return result
    End Function

    ''' <summary>
    ''' Sends a JSON-RPC POST for Streamable HTTP transport.
    ''' </summary>
    Private Async Function SendStreamableHTTPRequest(mcpUrl As String, payload As JObject, auth As MCPAuthInfo, Optional expectResponse As Boolean = True) As Task(Of JObject)
        EnsureTls12()

        Using handler As New HttpClientHandler()
            ConfigureHandler(handler)
            Using client As New HttpClient(handler)
                client.Timeout = TimeSpan.FromMilliseconds(MCP_DEFAULT_TIMEOUT)
                SetAcceptHeaders(client)
                SetAuthHeaders(client, auth)

                Dim content As New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                Dim response = Await client.PostAsync(mcpUrl, content).ConfigureAwait(False)

                If Not expectResponse Then
                    Try : Await response.Content.ReadAsStringAsync().ConfigureAwait(False) : Catch : End Try
                    Return New JObject()
                End If

                Dim responseText = Await response.Content.ReadAsStringAsync().ConfigureAwait(False)

                If Not response.IsSuccessStatusCode Then
                    Throw New Exception($"HTTP {CInt(response.StatusCode)} ({response.ReasonPhrase}): {If(responseText.Length > 500, responseText.Substring(0, 500) & "...", responseText)}")
                End If

                responseText = ExtractJsonFromSSE(responseText, response.Content.Headers.ContentType?.MediaType)

                If String.IsNullOrWhiteSpace(responseText) Then Return New JObject()

                Dim result = JObject.Parse(responseText)
                CheckJsonRpcError(result)
                Return result
            End Using
        End Using
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  SSE transport
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Initialises via SSE transport:
    '''   1. GET the SSE URL → receive "event: endpoint" with the POST URL
    '''   2. POST initialize to that endpoint
    '''   3. Read the initialize response from the SSE stream
    ''' </summary>
    Private Async Function MCPInitializeSSE(sseUrl As String, auth As MCPAuthInfo) As Task(Of MCPInitResult)
        ' Clean up any prior session
        If _mcpSseSession IsNot Nothing Then
            _mcpSseSession.Dispose()
            _mcpSseSession = Nothing
        End If

        EnsureTls12()

        Dim handler As New HttpClientHandler()
        ConfigureHandler(handler)
        Dim client As New HttpClient(handler)
        client.Timeout = TimeSpan.FromMilliseconds(MCP_DEFAULT_TIMEOUT)
        SetAuthHeaders(client, auth)
        client.DefaultRequestHeaders.Accept.Clear()
        client.DefaultRequestHeaders.Accept.Add(
            New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))

        ' Step 1: Open the SSE stream via GET
        Dim request As New HttpRequestMessage(HttpMethod.Get, sseUrl)
        Dim sseResponse = Await client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead).ConfigureAwait(False)

        If Not sseResponse.IsSuccessStatusCode Then
            Dim preview = Await sseResponse.Content.ReadAsStringAsync().ConfigureAwait(False)
            client.Dispose()
            handler.Dispose()
            Throw New Exception($"SSE GET {CInt(sseResponse.StatusCode)} ({sseResponse.ReasonPhrase}): {If(preview.Length > 300, preview.Substring(0, 300), preview)}")
        End If

        Dim stream = Await sseResponse.Content.ReadAsStreamAsync().ConfigureAwait(False)
        Dim reader As New StreamReader(stream, Encoding.UTF8)

        ' Step 2: Read SSE lines until we get "event: endpoint" + "data: ..."
        Dim postEndpoint As String = ""
        Dim currentEvent As String = ""
        Dim deadline = DateTime.UtcNow.AddMilliseconds(MCP_DEFAULT_TIMEOUT)

        While DateTime.UtcNow < deadline
            Dim line = Await ReadLineWithTimeout(reader, 15000).ConfigureAwait(False)
            If line Is Nothing Then Exit While

            If line.StartsWith("event:", StringComparison.OrdinalIgnoreCase) Then
                currentEvent = line.Substring(6).Trim()
            ElseIf line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                Dim data = line.Substring(5).Trim()
                If currentEvent.Equals("endpoint", StringComparison.OrdinalIgnoreCase) AndAlso Not String.IsNullOrWhiteSpace(data) Then
                    postEndpoint = data
                    Exit While
                End If
            End If
        End While

        If String.IsNullOrWhiteSpace(postEndpoint) Then
            reader.Dispose()
            sseResponse.Dispose()
            client.Dispose()
            Throw New Exception($"SSE stream at {sseUrl} did not return an endpoint event.")
        End If

        ' Resolve relative endpoint URL against the SSE base
        If postEndpoint.StartsWith("/") Then
            Dim uri As New Uri(sseUrl)
            postEndpoint = $"{uri.Scheme}://{uri.Authority}{postEndpoint}"
        ElseIf Not postEndpoint.StartsWith("http", StringComparison.OrdinalIgnoreCase) Then
            postEndpoint = sseUrl.TrimEnd("/"c) & "/" & postEndpoint
        End If

        ' Store the session
        _mcpSseSession = New MCPSSESession() With {
            .SseUrl = sseUrl,
            .PostEndpoint = postEndpoint,
            .Client = client,
            .StreamReader = reader,
            .SseResponse = sseResponse,
            .IsConnected = True
        }

        ' Step 3: POST initialize to the session endpoint
        Dim initPayload As New JObject From {
            {"jsonrpc", "2.0"}, {"id", 1}, {"method", "initialize"},
            {"params", New JObject From {
                {"protocolVersion", MCP_PROTOCOL_VERSION},
                {"capabilities", New JObject()},
                {"clientInfo", New JObject From {{"name", AN}, {"version", "1.0"}}}
            }}
        }

        Dim initResponse = Await SendSSERequest(initPayload, auth)

        ' For the INI, store the POST endpoint *without* the ephemeral session query string
        ' as a fallback, plus the SSE base URL for runtime session acquisition.
        Dim stablePostUrl As String = MCPStripQueryString(postEndpoint)

        Dim result As New MCPInitResult() With {
            .ResolvedUrl = stablePostUrl,
            .SseBaseUrl = sseUrl
        }
        Dim serverInfo = initResponse.SelectToken("result.serverInfo")
        If serverInfo IsNot Nothing Then
            result.ServerName = If(serverInfo("name")?.ToString(), "")
            result.ProtocolVersion = If(initResponse.SelectToken("result.protocolVersion")?.ToString(), "")
        End If

        ' Send initialized notification
        Dim notif As New JObject From {{"jsonrpc", "2.0"}, {"method", "notifications/initialized"}}
        Try
            Await SendSSERequest(notif, auth, expectResponse:=False)
        Catch
        End Try

        Return result
    End Function


    ''' <summary>
    ''' Sends a JSON-RPC request over an active SSE session:
    ''' POST to the session endpoint, then read the response from the SSE GET stream.
    ''' </summary>
    Private Async Function SendSSERequest(payload As JObject, auth As MCPAuthInfo, Optional expectResponse As Boolean = True) As Task(Of JObject)
        If _mcpSseSession Is Nothing OrElse Not _mcpSseSession.IsConnected Then
            Throw New Exception("No active SSE session.")
        End If

        ' POST the payload using a separate HttpClient (the session client is busy with GET)
        Using postHandler As New HttpClientHandler()
            ConfigureHandler(postHandler)
            Using postClient As New HttpClient(postHandler)
                postClient.Timeout = TimeSpan.FromMilliseconds(MCP_DEFAULT_TIMEOUT)
                SetAuthHeaders(postClient, auth)

                Dim content As New StringContent(payload.ToString(Formatting.None), Encoding.UTF8, "application/json")
                Dim postResponse = Await postClient.PostAsync(_mcpSseSession.PostEndpoint, content).ConfigureAwait(False)

                ' Some SSE servers return 200/202 on POST with empty body; the real response comes via SSE
                If Not postResponse.IsSuccessStatusCode Then
                    Dim errorText = Await postResponse.Content.ReadAsStringAsync().ConfigureAwait(False)
                    Throw New Exception($"SSE POST {CInt(postResponse.StatusCode)}: {If(errorText.Length > 300, errorText.Substring(0, 300), errorText)}")
                End If

                ' Check if the POST response itself contains a JSON-RPC result (some servers do this)
                Dim postBody = Await postResponse.Content.ReadAsStringAsync().ConfigureAwait(False)
                If Not String.IsNullOrWhiteSpace(postBody) AndAlso postBody.TrimStart().StartsWith("{") Then
                    Try
                        Dim directResult = JObject.Parse(postBody)
                        If directResult("result") IsNot Nothing OrElse directResult("error") IsNot Nothing Then
                            CheckJsonRpcError(directResult)
                            Return directResult
                        End If
                    Catch
                    End Try
                End If
            End Using
        End Using

        If Not expectResponse Then Return New JObject()

        ' Read the response from the SSE stream
        Dim expectedId As String = If(payload("id")?.ToString(), "")
        Dim deadline = DateTime.UtcNow.AddMilliseconds(MCP_DEFAULT_TIMEOUT)
        Dim currentEvent As String = ""

        While DateTime.UtcNow < deadline AndAlso _mcpSseSession.IsConnected
            Dim line = Await ReadLineWithTimeout(_mcpSseSession.StreamReader, 30000).ConfigureAwait(False)
            If line Is Nothing Then Exit While

            If line.StartsWith("event:", StringComparison.OrdinalIgnoreCase) Then
                currentEvent = line.Substring(6).Trim()
            ElseIf line.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                Dim data = line.Substring(5).Trim()
                If String.IsNullOrWhiteSpace(data) Then Continue While
                If currentEvent.Equals("endpoint", StringComparison.OrdinalIgnoreCase) Then Continue While

                ' Try to parse as JSON-RPC response
                If data.TrimStart().StartsWith("{") Then
                    Try
                        Dim msg = JObject.Parse(data)
                        ' Match by id if we have one
                        If Not String.IsNullOrWhiteSpace(expectedId) Then
                            Dim msgId = msg("id")?.ToString()
                            If msgId IsNot Nothing AndAlso msgId = expectedId Then
                                CheckJsonRpcError(msg)
                                Return msg
                            End If
                        End If
                        ' If no id matching needed (notification response) or id matches
                        If msg("result") IsNot Nothing OrElse msg("error") IsNot Nothing Then
                            CheckJsonRpcError(msg)
                            Return msg
                        End If
                    Catch
                    End Try
                End If
            End If
        End While

        Throw New Exception("Timeout waiting for SSE response.")
    End Function

    ''' <summary>Reads a line from the stream with a timeout.</summary>
    Private Shared Async Function ReadLineWithTimeout(reader As StreamReader, timeoutMs As Integer) As Task(Of String)
        Using cts As New CancellationTokenSource(timeoutMs)
            Try
                Dim lineTask = reader.ReadLineAsync()
                Dim completed = Await Task.WhenAny(lineTask, Task.Delay(timeoutMs, cts.Token)).ConfigureAwait(False)
                If completed Is lineTask Then
                    cts.Cancel()
                    Return Await lineTask
                End If
                Return Nothing
            Catch ex As OperationCanceledException
                Return Nothing
            Catch
                Return Nothing
            End Try
        End Using
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  Unified SendMCPRequest (routes to correct transport)
    ' ─────────────────────────────────────────────────────────────────────────

    ''' <summary>
    ''' Sends a JSON-RPC request using whichever transport is currently active.
    ''' If an SSE session exists, uses it; otherwise uses Streamable HTTP POST.
    ''' </summary>
    Private Async Function SendMCPRequest(mcpUrl As String, payload As JObject, auth As MCPAuthInfo, Optional expectResponse As Boolean = True) As Task(Of JObject)
        If _mcpSseSession IsNot Nothing AndAlso _mcpSseSession.IsConnected Then
            Return Await SendSSERequest(payload, auth, expectResponse)
        Else
            Return Await SendStreamableHTTPRequest(mcpUrl, payload, auth, expectResponse)
        End If
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  tools/list
    ' ─────────────────────────────────────────────────────────────────────────

    Private Async Function MCPListTools(mcpUrl As String, auth As MCPAuthInfo) As Task(Of List(Of MCPToolInfo))
        Dim tools As New List(Of MCPToolInfo)()
        Dim cursor As String = Nothing
        Dim requestId As Integer = 2

        Do
            Dim payload As New JObject From {
                {"jsonrpc", "2.0"}, {"id", requestId}, {"method", "tools/list"}
            }
            If cursor IsNot Nothing Then
                payload("params") = New JObject From {{"cursor", cursor}}
            End If

            Dim responseJson = Await SendMCPRequest(mcpUrl, payload, auth)

            Dim toolsArray = responseJson.SelectToken("result.tools")
            If toolsArray IsNot Nothing AndAlso toolsArray.Type = JTokenType.Array Then
                For Each t As JToken In toolsArray
                    Dim info As New MCPToolInfo()
                    info.Name = If(t("name")?.ToString(), "")
                    info.Description = If(t("description")?.ToString(), "")
                    Dim schema = t("inputSchema")
                    If schema IsNot Nothing AndAlso schema.Type = JTokenType.Object Then
                        info.InputSchema = CType(schema, JObject)
                    End If
                    tools.Add(info)
                Next
            End If

            cursor = responseJson.SelectToken("result.nextCursor")?.ToString()
            requestId += 1
        Loop While Not String.IsNullOrWhiteSpace(cursor)

        Return tools
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  Response key probing (Streamable HTTP only)
    ' ─────────────────────────────────────────────────────────────────────────

    Private Async Function ProbeMCPResponseKey(mcpUrl As String, auth As MCPAuthInfo, sampleTool As MCPToolInfo) As Task(Of String)
        Try
            Dim args As New JObject()
            If sampleTool.InputSchema IsNot Nothing Then
                Dim props = sampleTool.InputSchema("properties")
                Dim reqArray = sampleTool.InputSchema("required")
                Dim requiredNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
                If reqArray IsNot Nothing AndAlso reqArray.Type = JTokenType.Array Then
                    For Each r In reqArray : requiredNames.Add(r.ToString()) : Next
                End If
                If props IsNot Nothing AndAlso props.Type = JTokenType.Object Then
                    For Each prop In CType(props, JObject).Properties()
                        If requiredNames.Contains(prop.Name) Then
                            Dim pType = If(prop.Value("type")?.ToString(), "string")
                            Select Case pType
                                Case "string" : args(prop.Name) = "test"
                                Case "integer" : args(prop.Name) = If(prop.Value("minimum") IsNot Nothing, CInt(prop.Value("minimum")), 1)
                                Case "number" : args(prop.Name) = If(prop.Value("minimum") IsNot Nothing, CDbl(prop.Value("minimum")), 1.0)
                                Case "boolean" : args(prop.Name) = False
                                Case "array" : args(prop.Name) = New JArray()
                                Case "object" : args(prop.Name) = New JObject()
                                Case Else : args(prop.Name) = "test"
                            End Select
                        End If
                    Next
                End If
            End If

            Dim payload As New JObject From {
                {"jsonrpc", "2.0"}, {"id", 99}, {"method", "tools/call"},
                {"params", New JObject From {{"name", sampleTool.Name}, {"arguments", args}}}
            }
            Dim responseJson = Await SendStreamableHTTPRequest(mcpUrl, payload, auth)

            Dim contentArr = responseJson.SelectToken("result.content")
            If contentArr IsNot Nothing AndAlso contentArr.Type = JTokenType.Array AndAlso contentArr.HasValues Then
                Dim first = contentArr.First
                If first IsNot Nothing Then
                    If first("text") IsNot Nothing Then
                        If contentArr.Count() > 1 Then Return "{% for result.content %}{text}<CR>{% endfor %}"
                        Return "result.content[0].text"
                    End If
                    If first("data") IsNot Nothing Then Return "result.content[0].data"
                End If
                Return "result.content[0].text"
            End If
            If responseJson.SelectToken("result.output") IsNot Nothing Then Return "result.output"
            If responseJson.SelectToken("result.text") IsNot Nothing Then Return "result.text"
            If responseJson("response") IsNot Nothing Then Return "response"
            If responseJson("text") IsNot Nothing Then Return "text"
            Return "result.content[0].text"
        Catch
            Return "result.content[0].text"
        End Try
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  HTTP helpers
    ' ─────────────────────────────────────────────────────────────────────────

    Private Shared Sub EnsureTls12()
        If (System.Net.ServicePointManager.SecurityProtocol And System.Net.SecurityProtocolType.Tls12) = 0 Then
            System.Net.ServicePointManager.SecurityProtocol = System.Net.ServicePointManager.SecurityProtocol Or System.Net.SecurityProtocolType.Tls12
        End If
    End Sub

    Private Shared Sub ConfigureHandler(handler As HttpClientHandler)
        handler.AutomaticDecompression = System.Net.DecompressionMethods.GZip Or System.Net.DecompressionMethods.Deflate
        handler.UseProxy = True
        handler.Proxy = System.Net.WebRequest.DefaultWebProxy
        If handler.Proxy IsNot Nothing Then
            handler.Proxy.Credentials = System.Net.CredentialCache.DefaultCredentials
        End If
    End Sub

    Private Shared Sub SetAcceptHeaders(client As HttpClient)
        client.DefaultRequestHeaders.Accept.Clear()
        client.DefaultRequestHeaders.Accept.Add(New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"))
        client.DefaultRequestHeaders.Accept.Add(New System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("text/event-stream"))
    End Sub

    Private Shared Sub SetAuthHeaders(client As HttpClient, auth As MCPAuthInfo)
        If Not String.IsNullOrWhiteSpace(auth.HeaderA) AndAlso Not String.IsNullOrWhiteSpace(auth.HeaderB) Then
            Dim headerValue As String = auth.HeaderB.Replace("{apikey}", If(auth.LiveAPIKey, ""))
            client.DefaultRequestHeaders.TryAddWithoutValidation(auth.HeaderA, headerValue)
        End If
    End Sub

    Private Shared Sub CheckJsonRpcError(result As JObject)
        Dim err = result("error")
        If err IsNot Nothing Then
            Throw New Exception($"JSON-RPC error {If(err("code")?.ToString(), "")}: {If(err("message")?.ToString(), "Unknown error")}")
        End If
    End Sub

    ''' <summary>
    ''' Extracts the JSON-RPC message from an SSE-wrapped response body.
    ''' If the body is plain JSON, returns it unchanged.
    ''' </summary>
    Private Shared Function ExtractJsonFromSSE(responseText As String, contentType As String) As String
        Dim isSSE = (contentType IsNot Nothing AndAlso contentType.Equals("text/event-stream", StringComparison.OrdinalIgnoreCase)) OrElse
                    responseText.TrimStart().StartsWith("data:", StringComparison.OrdinalIgnoreCase)
        If Not isSSE Then Return responseText

        ' Find the last data: line that contains a JSON-RPC result or error
        Dim lastJson As String = Nothing
        For Each line In responseText.Split({vbCrLf, vbLf, vbCr}, StringSplitOptions.None)
            Dim trimmed = line.Trim()
            If trimmed.StartsWith("data:", StringComparison.OrdinalIgnoreCase) Then
                Dim data = trimmed.Substring(5).Trim()
                If data.Length > 0 AndAlso Not data.Equals("[DONE]", StringComparison.OrdinalIgnoreCase) AndAlso data.TrimStart().StartsWith("{") Then
                    Try
                        Dim candidate = JObject.Parse(data)
                        If candidate("result") IsNot Nothing OrElse candidate("error") IsNot Nothing Then
                            lastJson = data
                        End If
                    Catch
                    End Try
                End If
            End If
        Next
        Return If(lastJson, responseText)
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  Parameter extraction from JSON Schema
    ' ─────────────────────────────────────────────────────────────────────────

    Private Shared Function ExtractMCPParameters(schema As JObject) As List(Of MCPParamInfo)
        Dim result As New List(Of MCPParamInfo)()
        If schema Is Nothing Then Return result

        Dim props = schema("properties")
        If props Is Nothing OrElse props.Type <> JTokenType.Object Then Return result

        Dim reqArray = TryCast(schema("required"), JArray)
        Dim requiredNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
        If reqArray IsNot Nothing Then
            For Each r As JToken In reqArray
                Dim requiredName As String = If(r, "").ToString().Trim()
                If requiredName <> "" Then requiredNames.Add(requiredName)
            Next
        End If

        For Each prop As JProperty In CType(props, JObject).Properties()
            Dim propSchema As JObject = TryCast(prop.Value, JObject)
            If propSchema Is Nothing Then Continue For

            Dim p As New MCPParamInfo()
            p.Name = prop.Name
            p.IsRequired = requiredNames.Contains(prop.Name)
            p.Type = GetMCPCompatibleSchemaType(propSchema, schema)
            If String.IsNullOrWhiteSpace(p.Type) Then p.Type = "string"
            p.Description = If(propSchema("description")?.ToString(), prop.Name)

            Dim enumArr = TryCast(propSchema("enum"), JArray)
            If enumArr IsNot Nothing AndAlso enumArr.Count > 0 Then
                p.EnumValues = New List(Of String)()
                For Each e As JToken In enumArr
                    p.EnumValues.Add(e.ToString())
                Next
            End If

            Dim minToken = propSchema("minimum")
            If minToken IsNot Nothing Then
                Dim minV As Integer
                If Integer.TryParse(minToken.ToString(), minV) Then p.Minimum = minV
            End If

            Dim maxToken = propSchema("maximum")
            If maxToken IsNot Nothing Then
                Dim maxV As Integer
                If Integer.TryParse(maxToken.ToString(), maxV) Then p.Maximum = maxV
            End If

            Dim defToken = propSchema("default")
            If defToken IsNot Nothing Then
                p.DefaultValue = NormalizeMCPDefaultValue(defToken, p.Type)
            End If

            result.Add(p)
        Next

        Return result
    End Function

    Private Shared Function PrepareToolDefinitionSchema(inputSchema As JObject, normalizeSchema As Boolean) As JObject
        Dim baseSchema As JObject =
            If(TryCast(inputSchema?.DeepClone(), JObject),
               New JObject From {
                   {"type", "object"},
                   {"properties", New JObject()},
                   {"required", New JArray()}
               })

        If Not normalizeSchema Then
            Return CType(MCPSanitizeJsonStringValues(baseSchema), JObject)
        End If

        Dim normalized As JObject = TryCast(NormalizeMCPJsonSchema(baseSchema, baseSchema), JObject)

        If normalized Is Nothing Then
            normalized = New JObject From {
                {"type", "object"},
                {"properties", New JObject()},
                {"required", New JArray()}
            }
        End If

        If normalized("type") Is Nothing Then
            normalized("type") = "object"
        End If

        If String.Equals(If(normalized("type")?.ToString(), ""), "object", StringComparison.OrdinalIgnoreCase) AndAlso
           normalized("properties") Is Nothing Then
            normalized("properties") = New JObject()
        End If

        Return CType(MCPSanitizeJsonStringValues(normalized), JObject)
    End Function

    Private Shared Function NormalizeMCPJsonSchema(token As JToken, rootSchema As JObject) As JToken
        If token Is Nothing Then Return Nothing

        Dim resolvedToken As JToken = ResolveMCPReferenceToken(token, rootSchema)
        If resolvedToken Is Nothing Then Return Nothing

        Dim schemaObject As JObject = TryCast(resolvedToken, JObject)
        If schemaObject Is Nothing Then
            Return resolvedToken.DeepClone()
        End If

        Dim collapsed As JObject = CollapseMCPCompositeSchema(schemaObject, rootSchema)
        If collapsed IsNot Nothing AndAlso Not Object.ReferenceEquals(collapsed, schemaObject) Then
            Return NormalizeMCPJsonSchema(collapsed, rootSchema)
        End If

        Dim schemaType As String = GetMCPCompatibleSchemaType(schemaObject, rootSchema)
        If String.IsNullOrWhiteSpace(schemaType) Then schemaType = "string"

        Dim result As New JObject()

        Select Case schemaType
            Case "object"
                result("type") = "object"

                Dim normalizedProperties As New JObject()
                Dim propertiesToken As JObject = TryCast(ResolveMCPReferenceToken(schemaObject("properties"), rootSchema), JObject)

                If propertiesToken IsNot Nothing Then
                    For Each prop As JProperty In propertiesToken.Properties()
                        Dim normalizedProp As JObject = TryCast(NormalizeMCPJsonSchema(prop.Value, rootSchema), JObject)
                        If normalizedProp IsNot Nothing Then
                            normalizedProperties(prop.Name) = normalizedProp
                        End If
                    Next
                End If

                result("properties") = normalizedProperties

                Dim sourceRequired As JArray = TryCast(schemaObject("required"), JArray)
                If sourceRequired IsNot Nothing AndAlso sourceRequired.Count > 0 Then
                    Dim normalizedRequired As New JArray()
                    For Each req As JToken In sourceRequired
                        Dim reqName As String = If(req, "").ToString().Trim()
                        If reqName <> "" AndAlso normalizedProperties(reqName) IsNot Nothing Then
                            normalizedRequired.Add(reqName)
                        End If
                    Next

                    If normalizedRequired.Count > 0 Then
                        result("required") = normalizedRequired
                    End If
                End If

                Dim additionalPropertiesToken As JToken = schemaObject("additionalProperties")
                If additionalPropertiesToken IsNot Nothing AndAlso additionalPropertiesToken.Type = JTokenType.Boolean Then
                    result("additionalProperties") = additionalPropertiesToken.DeepClone()
                End If

            Case "array"
                result("type") = "array"

                Dim normalizedItems As JObject = TryCast(NormalizeMCPJsonSchema(schemaObject("items"), rootSchema), JObject)
                If normalizedItems Is Nothing Then
                    normalizedItems = New JObject From {{"type", "string"}}
                End If

                result("items") = normalizedItems

            Case Else
                result("type") = schemaType
        End Select

        ApplyMCPCompatibleSchemaMetadata(schemaObject, result)
        Return result
    End Function

    Private Shared Function CollapseMCPCompositeSchema(schemaObject As JObject, rootSchema As JObject) As JObject
        If schemaObject Is Nothing Then Return Nothing

        Dim alternatives As JArray = TryCast(schemaObject("anyOf"), JArray)
        If alternatives Is Nothing OrElse alternatives.Count = 0 Then
            alternatives = TryCast(schemaObject("oneOf"), JArray)
        End If

        If alternatives IsNot Nothing AndAlso alternatives.Count > 0 Then
            Dim bestMatch As JObject = Nothing
            Dim bestScore As Integer = Integer.MinValue

            For Each candidate As JToken In alternatives
                Dim candidateObject As JObject = TryCast(ResolveMCPReferenceToken(candidate, rootSchema), JObject)
                If candidateObject Is Nothing Then Continue For

                Dim candidateType As String = GetMCPCompatibleSchemaType(candidateObject, rootSchema)
                If String.IsNullOrWhiteSpace(candidateType) OrElse
                   String.Equals(candidateType, "null", StringComparison.OrdinalIgnoreCase) Then
                    Continue For
                End If

                Dim score As Integer = GetMCPTypePreferenceScore(candidateType)
                If score > bestScore Then
                    bestScore = score
                    bestMatch = CType(candidateObject.DeepClone(), JObject)
                End If
            Next

            If bestMatch IsNot Nothing Then
                If bestMatch("description") Is Nothing AndAlso schemaObject("description") IsNot Nothing Then
                    bestMatch("description") = schemaObject("description").DeepClone()
                End If
                Return bestMatch
            End If
        End If

        Return schemaObject
    End Function

    Private Shared Function GetMCPCompatibleSchemaType(schemaObject As JObject, rootSchema As JObject) As String
        If schemaObject Is Nothing Then Return "string"

        Dim typeToken As JToken = schemaObject("type")

        If typeToken IsNot Nothing Then
            If typeToken.Type = JTokenType.String Then
                Dim typeName As String = typeToken.ToString().Trim().ToLowerInvariant()
                If typeName <> "" AndAlso typeName <> "null" Then Return typeName
            ElseIf typeToken.Type = JTokenType.Array Then
                For Each item As JToken In CType(typeToken, JArray)
                    Dim typeName As String = item.ToString().Trim().ToLowerInvariant()
                    If typeName <> "" AndAlso typeName <> "null" Then Return typeName
                Next
            End If
        End If

        If schemaObject("properties") IsNot Nothing Then Return "object"
        If schemaObject("items") IsNot Nothing Then Return "array"

        Dim composite As JObject = CollapseMCPCompositeSchema(schemaObject, rootSchema)
        If composite IsNot Nothing AndAlso Not Object.ReferenceEquals(composite, schemaObject) Then
            Return GetMCPCompatibleSchemaType(composite, rootSchema)
        End If

        If schemaObject("enum") IsNot Nothing Then Return "string"

        Return "string"
    End Function

    Private Shared Function ResolveMCPReferenceToken(token As JToken, rootSchema As JObject) As JToken
        If token Is Nothing Then Return Nothing

        Dim obj As JObject = TryCast(token, JObject)
        If obj Is Nothing Then Return token

        Dim refValue As String = If(obj("$ref")?.ToString(), "").Trim()
        If String.IsNullOrWhiteSpace(refValue) Then
            Return obj
        End If

        Dim resolved As JToken = ResolveMCPJsonPointer(rootSchema, refValue)
        If resolved Is Nothing Then
            Return obj
        End If

        Dim resolvedObject As JObject = TryCast(resolved.DeepClone(), JObject)
        If resolvedObject Is Nothing Then
            Return resolved.DeepClone()
        End If

        For Each prop As JProperty In obj.Properties()
            If prop.Name.Equals("$ref", StringComparison.OrdinalIgnoreCase) Then Continue For
            resolvedObject(prop.Name) = prop.Value.DeepClone()
        Next

        Return resolvedObject
    End Function

    Private Shared Function ResolveMCPJsonPointer(rootToken As JToken, pointer As String) As JToken
        If rootToken Is Nothing OrElse String.IsNullOrWhiteSpace(pointer) Then Return Nothing
        If Not pointer.StartsWith("#/", StringComparison.Ordinal) Then Return Nothing

        Dim current As JToken = rootToken

        For Each rawSegment As String In pointer.Substring(2).Split("/"c)
            Dim segment As String = rawSegment.Replace("~1", "/").Replace("~0", "~")

            If current Is Nothing Then Return Nothing

            If current.Type = JTokenType.Object Then
                current = current(segment)
            ElseIf current.Type = JTokenType.Array Then
                Dim index As Integer
                If Not Integer.TryParse(segment, index) Then Return Nothing
                Dim arr As JArray = CType(current, JArray)
                If index < 0 OrElse index >= arr.Count Then Return Nothing
                current = arr(index)
            Else
                Return Nothing
            End If
        Next

        Return current
    End Function

    Private Shared Sub ApplyMCPCompatibleSchemaMetadata(source As JObject, target As JObject)
        If source Is Nothing OrElse target Is Nothing Then Return

        If source("description") IsNot Nothing Then
            target("description") = source("description").DeepClone()
        End If

        Dim targetType As String = If(target("type")?.ToString(), "").Trim().ToLowerInvariant()

        If targetType <> "object" AndAlso targetType <> "array" Then
            Dim enumToken As JArray = TryCast(source("enum"), JArray)
            If enumToken IsNot Nothing AndAlso enumToken.Count > 0 Then
                target("enum") = enumToken.DeepClone()
            End If
        End If

        If targetType = "integer" OrElse targetType = "number" Then
            If source("minimum") IsNot Nothing Then target("minimum") = source("minimum").DeepClone()
            If source("maximum") IsNot Nothing Then target("maximum") = source("maximum").DeepClone()
        End If
    End Sub

    Private Shared Function GetMCPTypePreferenceScore(typeName As String) As Integer
        Select Case If(typeName, "").Trim().ToLowerInvariant()
            Case "object" : Return 60
            Case "array" : Return 50
            Case "string" : Return 40
            Case "integer" : Return 30
            Case "number" : Return 20
            Case "boolean" : Return 10
            Case Else : Return 0
        End Select
    End Function

    Private Shared Function NormalizeMCPDefaultValue(defaultToken As JToken, parameterType As String) As String
        If defaultToken Is Nothing Then
            Return ""
        End If

        If defaultToken.Type = JTokenType.Null Then
            Return ""
        End If

        Select Case If(parameterType, "").Trim().ToLowerInvariant()
            Case "boolean"
                Dim boolValue As Boolean
                If Boolean.TryParse(defaultToken.ToString(), boolValue) Then
                    Return If(boolValue, "true", "false")
                End If
                Return "false"

            Case "integer"
                Dim longValue As Long
                If Long.TryParse(defaultToken.ToString(), Globalization.NumberStyles.Integer, Globalization.CultureInfo.InvariantCulture, longValue) OrElse
                   Long.TryParse(defaultToken.ToString(), longValue) Then
                    Return longValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If
                Return "0"

            Case "number"
                Dim doubleValue As Double
                Dim normalized As String = defaultToken.ToString().Replace(","c, "."c)

                If Double.TryParse(normalized, Globalization.NumberStyles.Float Or Globalization.NumberStyles.AllowThousands,
                                   Globalization.CultureInfo.InvariantCulture, doubleValue) OrElse
                   Double.TryParse(defaultToken.ToString(), doubleValue) Then
                    Return doubleValue.ToString(Globalization.CultureInfo.InvariantCulture)
                End If

                Return "0"

            Case "array", "object"
                Return defaultToken.ToString(Formatting.None)

            Case Else
                Return defaultToken.ToString()
        End Select
    End Function

    Private Shared Function BuildToolParameterDefaultValue(p As MCPParamInfo) As String
        Dim defVal As String = If(p.DefaultValue, "")

        If Not String.IsNullOrWhiteSpace(defVal) Then
            Return defVal
        End If

        If Not p.IsRequired Then
            Return ""
        End If

        Select Case p.Type
            Case "integer", "number"
                Return If(p.Minimum.HasValue, p.Minimum.Value.ToString(Globalization.CultureInfo.InvariantCulture), "0")

            Case "boolean"
                Return "false"

            Case "array"
                Return "[]"

            Case "object"
                Return "{}"

            Case Else
                If p.EnumValues IsNot Nothing AndAlso p.EnumValues.Count > 0 Then
                    Return p.EnumValues(0)
                End If

                Return ""
        End Select
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  INI section builder
    ' ─────────────────────────────────────────────────────────────────────────

    Private Function BuildINISectionForTool(tool As MCPToolInfo, sectionName As String, mcpUrl As String,
                                            sseBaseUrl As String, auth As MCPAuthInfo, responseKey As String,
                                            mergePrompt As String, timeout As Long,
                                            warnings As List(Of String), toolOnly As Boolean,
                                            normalizeSchema As Boolean) As String
        Dim sb As New StringBuilder()
        Dim effectiveSchema As JObject = PrepareToolDefinitionSchema(tool.InputSchema, normalizeSchema)

        Dim parameterSchema As JObject =
            If(TryCast(tool.InputSchema?.DeepClone(), JObject), effectiveSchema)

        Dim allParams = ExtractMCPParameters(parameterSchema)

        Dim primaryParam As MCPParamInfo = Nothing
        For Each p In allParams
            If p.IsRequired AndAlso p.Type = "string" Then primaryParam = p : Exit For
        Next
        If primaryParam Is Nothing Then
            For Each p In allParams
                If p.IsRequired Then primaryParam = p : Exit For
            Next
        End If

        Dim remainingParams As New List(Of MCPParamInfo)()
        For Each p In allParams
            If primaryParam IsNot Nothing AndAlso p.Name = primaryParam.Name Then Continue For
            remainingParams.Add(p)
        Next

        If remainingParams.Count > MCP_MAX_INI_PARAMS Then
            warnings.Add(
                $"'{sectionName}': {remainingParams.Count} extra parameters. " &
                $"Mapped 1–4: {String.Join(", ", remainingParams.Take(MCP_MAX_INI_PARAMS).Select(Function(x) x.Name))}. " &
                $"Overflow (ToolAPICall only): {String.Join(", ", remainingParams.Skip(MCP_MAX_INI_PARAMS).Select(Function(x) x.Name))}")
        End If

        Dim iniParams As New List(Of MCPParamInfo)()
        For i As Integer = 0 To Math.Min(MCP_MAX_INI_PARAMS, remainingParams.Count) - 1
            iniParams.Add(remainingParams(i))
        Next

        Dim endpointValue As String
        If Not String.IsNullOrWhiteSpace(sseBaseUrl) Then
            endpointValue = MCP_SSE_PREFIX & sseBaseUrl
        Else
            endpointValue = SharedMethods.MCP_STREAMABLE_PREFIX & mcpUrl
        End If

        sb.AppendLine($"[{MCPSanitizeINIValue(sectionName)}]")
        sb.AppendLine()
        If Not String.IsNullOrWhiteSpace(tool.Description) Then
            sb.AppendLine($"; {tool.Description.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")}")
            sb.AppendLine()
        End If

        sb.AppendLine($"APIKey = {MCPSanitizeINIValue(auth.APIKeyPlaceholder)}")
        sb.AppendLine($"APIKeyPrefix = {MCPSanitizeINIValue(auth.APIKeyPrefix)}")
        sb.AppendLine($"APIKeyEncrypted = {auth.APIKeyEncrypted}")

        If auth.OAuth2 Then
            Dim oauthEndpointValue As String

            If auth.OAuth2UseDeviceFlow Then
                oauthEndpointValue = "device:" & auth.OAuth2DeviceAuthorizationEndpoint & "¦" & auth.OAuth2TokenEndpoint
            Else
                oauthEndpointValue = auth.OAuth2AuthorizationEndpoint & "¦" & auth.OAuth2TokenEndpoint
            End If

            If Not String.IsNullOrWhiteSpace(auth.OAuth2Resource) Then
                oauthEndpointValue &= "¦" & auth.OAuth2Resource
            End If

            sb.AppendLine("OAuth2 = True")
            sb.AppendLine($"OAuth2ClientMail = {MCPSanitizeINIValue(auth.OAuth2ClientId)}")
            sb.AppendLine($"OAuth2Scopes = {MCPSanitizeINIValue(auth.OAuth2Scope)}")
            sb.AppendLine($"OAuth2Endpoint = {MCPSanitizeINIValue(oauthEndpointValue)}")
            sb.AppendLine($"OAuth2ATExpiry = {auth.OAuth2ExpiresIn}")
        End If

        sb.AppendLine($"Model = {MCPSanitizeINIValue(sectionName)}")
        sb.AppendLine($"Endpoint = {MCPSanitizeINIValue(endpointValue)}")
        sb.AppendLine($"HeaderA = {MCPSanitizeINIValue(auth.HeaderA)}")
        sb.AppendLine($"HeaderB = {MCPSanitizeINIValue(auth.HeaderB)}")
        sb.AppendLine($"Response = {MCPSanitizeINIValue(responseKey)}")
        sb.AppendLine($"Timeout = {timeout}")

        Dim apiCallJson As New JObject From {
            {"jsonrpc", "2.0"}, {"id", 0}, {"method", "tools/call"},
            {"params", New JObject From {
                {"name", tool.Name},
                {"arguments", BuildAPICallArguments(primaryParam, iniParams, allParams, True)}
            }}
        }
        sb.AppendLine($"APICall = {apiCallJson.ToString(Formatting.None)}")

        For i As Integer = 0 To iniParams.Count - 1
            sb.AppendLine($"Parameter{i + 1} = {FormatINIParameter(iniParams(i))}")
        Next

        If Not String.IsNullOrWhiteSpace(mergePrompt) Then sb.AppendLine($"MergePrompt = {MCPSanitizeINIValue(mergePrompt)}")

        sb.AppendLine($"Tool = True")
        sb.AppendLine($"ToolOnly = {toolOnly}")
        sb.AppendLine($"ToolName = {MCPSanitizeINIValue(tool.Name)}")
        sb.AppendLine($"ToolPriority = 10")
        sb.AppendLine($"ToolErrorHandling = skip")

        Dim toolApiCallJson As New JObject From {
            {"jsonrpc", "2.0"}, {"id", 0}, {"method", "tools/call"},
            {"params", New JObject From {
                {"name", tool.Name},
                {"arguments", BuildToolAPICallArguments(allParams)}
            }}
        }
        sb.AppendLine($"ToolAPICall = {toolApiCallJson.ToString(Formatting.None)}")

        Dim defaults As New JObject()
        For Each p In allParams
            If Not p.IsRequired Then
                defaults(p.Name) = BuildToolParameterDefaultValue(p)
            End If
        Next
        sb.AppendLine($"ToolParameterDefaults = {defaults.ToString(Formatting.None)}")

        Dim safeDescription As String = If(tool.Description, "").Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
        Dim instrPrompt As New StringBuilder()
        instrPrompt.Append($"{tool.Name}: {safeDescription} Parameters: ")
        For Each p In allParams
            Dim safeParamDesc As String = If(p.Description, "").Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
            instrPrompt.Append($"{p.Name} ({p.Type}, {If(p.IsRequired, "required", "optional")})")
            If Not String.IsNullOrWhiteSpace(safeParamDesc) Then instrPrompt.Append($" - {safeParamDesc}")
            If p.EnumValues IsNot Nothing AndAlso p.EnumValues.Count > 0 Then instrPrompt.Append($" [values: {String.Join(", ", p.EnumValues)}]")
            instrPrompt.Append(". ")
        Next
        sb.AppendLine($"ToolInstructionsPrompt = {instrPrompt.ToString().Trim()}")

        Dim safeInputSchema As JObject =
            If(effectiveSchema,
               New JObject From {
                   {"type", "object"},
                   {"properties", New JObject()},
                   {"required", New JArray()}
               })

        Dim toolDef As New JObject From {
            {"name", tool.Name},
            {"description", safeDescription},
            {"parameters", safeInputSchema}
        }
        sb.AppendLine($"ToolDefinition = {toolDef.ToString(Formatting.None)}")
        sb.AppendLine()
        Return sb.ToString()
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  APICall argument builders
    ' ─────────────────────────────────────────────────────────────────────────

    Private Shared Function BuildAPICallArguments(primaryParam As MCPParamInfo, iniParams As List(Of MCPParamInfo),
                                                  allParams As List(Of MCPParamInfo), usePromptUser As Boolean) As JObject
        Dim args As New JObject()
        For Each p In allParams
            If primaryParam IsNot Nothing AndAlso p.Name = primaryParam.Name Then
                args(p.Name) = If(usePromptUser, "{promptuser}", "{query}")
            Else
                Dim iniIdx As Integer = -1
                For i As Integer = 0 To iniParams.Count - 1
                    If iniParams(i).Name = p.Name Then iniIdx = i : Exit For
                Next
                If iniIdx >= 0 Then
                    Dim placeholder = "{parameter" & (iniIdx + 1) & "}"
                    Select Case p.Type
                        Case "integer", "number", "array", "boolean"
                            args.Add(New JProperty(p.Name, New JRaw(placeholder)))
                        Case Else
                            args(p.Name) = placeholder
                    End Select
                Else
                    EmitDefaultArgument(args, p)
                End If
            End If
        Next
        Return args
    End Function

    Private Shared Function BuildToolAPICallArguments(allParams As List(Of MCPParamInfo)) As JObject
        Dim args As New JObject()
        For Each p In allParams
            Dim placeholder = "{" & p.Name & "}"
            Select Case p.Type
                Case "integer", "number", "boolean", "array"
                    args.Add(New JProperty(p.Name, New JRaw(placeholder)))
                Case Else
                    args(p.Name) = placeholder
            End Select
        Next
        Return args
    End Function

    Private Shared Sub EmitDefaultArgument(args As JObject, p As MCPParamInfo)
        Dim defVal = BuildToolParameterDefaultValue(p)

        If String.IsNullOrWhiteSpace(defVal) AndAlso Not p.IsRequired Then
            Return
        End If

        Select Case p.Type
            Case "integer"
                Dim iv As Integer = 0
                If Not String.IsNullOrEmpty(defVal) Then Integer.TryParse(defVal, iv)
                If iv = 0 AndAlso p.Minimum.HasValue Then iv = p.Minimum.Value
                args.Add(New JProperty(p.Name, New JRaw(iv.ToString())))

            Case "number"
                Dim dv As Double = 0
                If Not String.IsNullOrEmpty(defVal) Then
                    Double.TryParse(defVal.Replace(","c, "."c),
                                    Globalization.NumberStyles.Float,
                                    Globalization.CultureInfo.InvariantCulture,
                                    dv)
                End If
                If dv = 0 AndAlso p.Minimum.HasValue Then dv = p.Minimum.Value
                args.Add(New JProperty(p.Name, New JRaw(dv.ToString("0.###", Globalization.CultureInfo.InvariantCulture))))

            Case "boolean"
                Dim bv As Boolean = False
                If Not String.IsNullOrEmpty(defVal) Then Boolean.TryParse(defVal, bv)
                args.Add(New JProperty(p.Name, New JRaw(If(bv, "true", "false"))))

            Case "array"
                If Not String.IsNullOrEmpty(defVal) AndAlso defVal.TrimStart().StartsWith("[") Then
                    args.Add(New JProperty(p.Name, New JRaw(defVal)))
                ElseIf p.IsRequired Then
                    args.Add(New JProperty(p.Name, New JRaw("[]")))
                End If

            Case "object"
                If Not String.IsNullOrEmpty(defVal) AndAlso defVal.TrimStart().StartsWith("{") Then
                    args.Add(New JProperty(p.Name, New JRaw(defVal)))
                ElseIf p.IsRequired Then
                    args.Add(New JProperty(p.Name, New JRaw("{}")))
                End If

            Case Else
                If Not String.IsNullOrWhiteSpace(defVal) OrElse p.IsRequired Then
                    args(p.Name) = If(defVal, "")
                End If
        End Select
    End Sub

    ' ─────────────────────────────────────────────────────────────────────────
    '  INI Parameter formatting
    ' ─────────────────────────────────────────────────────────────────────────

    Private Shared Function FormatINIParameter(p As MCPParamInfo) As String
        Dim sb As New StringBuilder()
        Dim desc = If(Not String.IsNullOrWhiteSpace(p.Description), p.Description, p.Name)
        ' Ensure single-line: strip CR/LF and semicolons
        desc = desc.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ").Replace(";"c, ","c)
        If desc.Length > 80 Then desc = desc.Substring(0, 77) & "..."
        sb.Append(desc)

        Dim iniType As String
        Select Case p.Type
            Case "integer" : iniType = "Integer"
            Case "number" : iniType = "Double"
            Case "boolean" : iniType = "Boolean"
            Case Else : iniType = "String"
        End Select
        sb.Append("; " & iniType)

        Dim defVal = p.DefaultValue
        If String.IsNullOrEmpty(defVal) Then
            If p.EnumValues IsNot Nothing AndAlso p.EnumValues.Count > 0 Then
                defVal = p.EnumValues(0)
            ElseIf p.Type = "boolean" Then
                defVal = "false"
            ElseIf (p.Type = "integer" OrElse p.Type = "number") AndAlso p.Minimum.HasValue Then
                defVal = p.Minimum.Value.ToString()
            Else
                defVal = ""
            End If
        End If
        sb.Append("; " & defVal)

        If (iniType = "Integer" OrElse iniType = "Double") AndAlso p.Minimum.HasValue AndAlso p.Maximum.HasValue Then
            sb.Append($"; {p.Minimum.Value}-{p.Maximum.Value}")
        End If

        If p.EnumValues IsNot Nothing AndAlso p.EnumValues.Count > 0 AndAlso iniType = "String" Then
            sb.Append("; ")
            sb.Append(String.Join(", ", p.EnumValues.Select(Function(o) o.Replace(",", "\\,"))))
        End If
        Return sb.ToString()
    End Function

    ' ─────────────────────────────────────────────────────────────────────────
    '  Utilities
    ' ─────────────────────────────────────────────────────────────────────────

    Private Shared Function MCPCapitalizeFirst(s As String) As String
        If String.IsNullOrEmpty(s) Then Return s
        Return Char.ToUpperInvariant(s(0)) & s.Substring(1)
    End Function

    Private Shared Function MCPSanitizeFileName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "mcp_server"
        Dim invalid = Path.GetInvalidFileNameChars()
        Dim result As New StringBuilder(name.Length)
        For Each c In name
            If Array.IndexOf(invalid, c) >= 0 Then result.Append("_"c) Else result.Append(c)
        Next
        Return result.ToString().Trim().Trim("."c)
    End Function

    ''' <summary>
    ''' Converts a snake_case or camelCase tool name into a human-friendly title.
    ''' e.g. "search_case_law" → "Search Case Law", "getDecision" → "Get Decision"
    ''' </summary>
    Private Shared Function MCPHumanizeName(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return name

        ' Replace underscores and hyphens with spaces
        Dim spaced = name.Replace("_"c, " "c).Replace("-"c, " "c)

        ' Insert space before each uppercase letter that follows a lowercase (camelCase split)
        Dim sb As New StringBuilder()
        For i As Integer = 0 To spaced.Length - 1
            Dim c = spaced(i)
            If i > 0 AndAlso Char.IsUpper(c) AndAlso Char.IsLower(spaced(i - 1)) Then
                sb.Append(" "c)
            End If
            sb.Append(c)
        Next

        ' Title-case each word
        Dim words = sb.ToString().Split({" "c}, StringSplitOptions.RemoveEmptyEntries)
        For w As Integer = 0 To words.Length - 1
            If words(w).Length > 0 Then
                words(w) = Char.ToUpperInvariant(words(w)(0)) & words(w).Substring(1).ToLowerInvariant()
            End If
        Next
        Return String.Join(" ", words)
    End Function

    ''' <summary>
    ''' Strips CR/LF characters from a string value so it remains on a single INI line.
    ''' </summary>
    Private Shared Function MCPSanitizeINIValue(value As String) As String
        If String.IsNullOrEmpty(value) Then Return value
        Return value.Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " ")
    End Function

    ''' <summary>
    ''' Deep-clones a JObject, replacing any CR/LF in string values with spaces.
    ''' Ensures the serialised JSON stays on one line when written to an INI file.
    ''' </summary>
    Private Shared Function MCPSanitizeJsonStringValues(token As JToken) As JToken
        If token Is Nothing Then Return token
        Select Case token.Type
            Case JTokenType.Object
                Dim obj As New JObject()
                For Each prop In CType(token, JObject).Properties()
                    obj(prop.Name) = MCPSanitizeJsonStringValues(prop.Value)
                Next
                Return obj
            Case JTokenType.Array
                Dim arr As New JArray()
                For Each item In CType(token, JArray)
                    arr.Add(MCPSanitizeJsonStringValues(item))
                Next
                Return arr
            Case JTokenType.String
                Return New JValue(token.ToString().Replace(vbCrLf, " ").Replace(vbLf, " ").Replace(vbCr, " "))
            Case Else
                Return token.DeepClone()
        End Select
    End Function

    ''' <summary>
    ''' Removes the query string from a URL, returning only the scheme + authority + path.
    ''' Used to strip ephemeral session IDs from SSE transport POST endpoints.
    ''' e.g. "https://server/messages?sessionId=abc" → "https://server/messages"
    ''' </summary>
    Private Shared Function MCPStripQueryString(url As String) As String
        If String.IsNullOrWhiteSpace(url) Then Return url
        Try
            Dim uri As New Uri(url)
            Return uri.GetLeftPart(UriPartial.Path)
        Catch
            ' Fallback: strip everything from '?' onwards
            Dim qIdx = url.IndexOf("?"c)
            If qIdx > 0 Then Return url.Substring(0, qIdx)
            Return url
        End Try
    End Function


End Class
