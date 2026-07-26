' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: PythonExecuteTool.vb
' Purpose: Exposes the shared `python_execute` tool as a Red Ink `ModelConfig`
'          and bridges host callers to the host-agnostic execution core.
'
' Architecture / How it works:
'  - Maintains one configured `PythonExecuteToolCoreOptions` instance for the
'    process and validates availability before registration or execution.
'  - Builds the advertised tool metadata (`ToolName`, instructions, schema, and
'    model description) from `PythonExecuteToolCore`.
'  - Executes requests via `ExecuteAsync()` / `ExecuteDetailedAsync()`, with
'    optional per-call overrides for allowed host operations and host-service
'    handlers without mutating shared global configuration.
'  - Produces consistent host-failure payloads for callers that need a safe JSON
'    envelope even when execution cannot start or is cancelled.
' =============================================================================


Option Explicit On
Option Strict On
Option Infer On
Imports SharedLibrary.SharedLibrary

Namespace Agents

    ' Host adapter for the existing Red Ink agent infrastructure.
    ' This file is compiled in the target SharedLibrary project together with
    ' RedInkPythonAgentClient.vb and PythonExecuteToolCore.vb.
    Public NotInheritable Class PythonExecuteTool
        Public Const ToolName As System.String = "python_execute"

        Public Shared Function IsPythonTool(name As System.String) As System.Boolean
            Return Not System.String.IsNullOrWhiteSpace(name) AndAlso
                   System.String.Equals(name, ToolName, System.StringComparison.OrdinalIgnoreCase)
        End Function

        Private Shared ReadOnly ConfigurationLock As New System.Object()
        Private Shared ConfiguredOptions As Agents.PythonExecuteToolCoreOptions

        Private Sub New()
        End Sub

        Public Shared Sub Configure(options As Agents.PythonExecuteToolCoreOptions)
            If options Is Nothing Then
                Throw New System.ArgumentNullException(NameOf(options))
            End If
            Agents.PythonExecuteToolCore.ResolveAndValidateAvailability(options)
            SyncLock ConfigurationLock
                ConfiguredOptions = options
            End SyncLock
        End Sub

        Public Shared Function IsAvailable(ByRef errorCode As System.String) As System.Boolean
            Dim options As Agents.PythonExecuteToolCoreOptions = GetConfiguredOptionsOrNothing()
            If options Is Nothing Then
                errorCode = "CONFIGURATION_INVALID"
                Return False
            End If
            Try
                Agents.PythonExecuteToolCore.ResolveAndValidateAvailability(options)
                errorCode = System.String.Empty
                Return True
            Catch ex As RedInkPythonAgentExecutableTrustException
                errorCode = ex.Code
                Return False
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                errorCode = "CONFIGURATION_INVALID"
                Return False
            End Try
        End Function

        Public Shared Function Build(
            context As System.Object,
            Optional toolPriority As System.Int32 = 996,
            Optional displaySuffix As System.String = ""
        ) As ModelConfig
            Dim ignoredErrorCode As System.String = System.String.Empty
            If Not IsAvailable(ignoredErrorCode) Then
                Throw New RedInkPythonAgentConfigurationException("The python_execute tool is unavailable: " & ignoredErrorCode)
            End If
            Return New ModelConfig() With {
                .ToolName = ToolName,
                .ToolInstructionsPrompt = Agents.PythonExecuteToolCore.ToolInstructionsPrompt,
                .ToolDefinition = Agents.PythonExecuteToolCore.ToolDefinitionJson,
                .ModelDescription = "Secure Python Execution" & displaySuffix,
                .Tool = True,
                .ToolPriority = toolPriority,
                .ToolErrorHandling = "skip"
            }
        End Function

        Public Shared Function TryBuild(
            context As System.Object,
            ByRef modelConfig As ModelConfig,
            Optional toolPriority As System.Int32 = 996,
            Optional displaySuffix As System.String = ""
        ) As System.Boolean
            Dim errorCode As System.String = System.String.Empty
            If Not IsAvailable(errorCode) Then
                modelConfig = Nothing
                Return False
            End If
            modelConfig = Build(context, toolPriority, displaySuffix)
            Return True
        End Function

        Public Shared Async Function ExecuteAsync(
            context As System.Object,
            arguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken,
            Optional logStep As System.Action(Of System.String) = Nothing,
            Optional logInfo As System.Action(Of System.String) = Nothing,
            Optional logWarn As System.Action(Of System.String) = Nothing,
            Optional logDiag As System.Action(Of System.String) = Nothing,
            Optional hostServiceHandler As IRedInkPythonAgentHostServiceHandler = Nothing,
            Optional allowedOperations As System.Collections.Generic.IEnumerable(Of System.String) = Nothing
        ) As System.Threading.Tasks.Task(Of System.String)
            Dim result As PythonExecuteToolCoreResult = Await ExecuteDetailedAsync(context, arguments, cancellationToken, logStep, logInfo, logWarn, logDiag, hostServiceHandler, allowedOperations).ConfigureAwait(False)
            Return result.Payload
        End Function

        ''' <summary>
        ''' Executes python_execute. When <paramref name="hostServiceHandler"/> and
        ''' <paramref name="allowedOperations"/> are supplied, they are applied for this call only
        ''' (via a cloned options instance), scoping host-mediated capabilities to what the current
        ''' tooling loop exposes. When omitted, the globally configured options are used unchanged,
        ''' preserving existing behavior for all current callers.
        ''' </summary>
        Public Shared Async Function ExecuteDetailedAsync(
            context As System.Object,
            arguments As System.Collections.Generic.Dictionary(Of System.String, System.Object),
            cancellationToken As System.Threading.CancellationToken,
            Optional logStep As System.Action(Of System.String) = Nothing,
            Optional logInfo As System.Action(Of System.String) = Nothing,
            Optional logWarn As System.Action(Of System.String) = Nothing,
            Optional logDiag As System.Action(Of System.String) = Nothing,
            Optional hostServiceHandler As IRedInkPythonAgentHostServiceHandler = Nothing,
            Optional allowedOperations As System.Collections.Generic.IEnumerable(Of System.String) = Nothing
        ) As System.Threading.Tasks.Task(Of PythonExecuteToolCoreResult)
            Dim options As PythonExecuteToolCoreOptions = GetConfiguredOptions()
            If hostServiceHandler IsNot Nothing OrElse allowedOperations IsNot Nothing Then
                options = options.Clone()
                options.HostServiceHandler = hostServiceHandler
                options.AllowedOperations = If(allowedOperations Is Nothing,
                                               New System.Collections.Generic.List(Of System.String)(),
                                               New System.Collections.Generic.List(Of System.String)(allowedOperations))
            End If
            Return Await PythonExecuteToolCore.ExecuteAsync(options, arguments, cancellationToken, logStep, logInfo, logWarn, logDiag).ConfigureAwait(False)
        End Function


        Public Shared Function CreateHostFailurePayload(status As System.String, code As System.String) As System.String
            Return New Newtonsoft.Json.Linq.JObject(
                New Newtonsoft.Json.Linq.JProperty("status", status),
                New Newtonsoft.Json.Linq.JProperty("exit_code", 1),
                New Newtonsoft.Json.Linq.JProperty("duration_ms", 0),
                New Newtonsoft.Json.Linq.JProperty("diagnostic_id", System.Guid.NewGuid().ToString("D")),
                New Newtonsoft.Json.Linq.JProperty("human_log_available", False),
                New Newtonsoft.Json.Linq.JProperty("result", Newtonsoft.Json.Linq.JValue.CreateNull()),
                New Newtonsoft.Json.Linq.JProperty("output_files", New Newtonsoft.Json.Linq.JArray()),
                New Newtonsoft.Json.Linq.JProperty("error", New Newtonsoft.Json.Linq.JObject(
                    New Newtonsoft.Json.Linq.JProperty("code", code),
                    New Newtonsoft.Json.Linq.JProperty("phase", "initializing"),
                    New Newtonsoft.Json.Linq.JProperty("retryable", False),
                    New Newtonsoft.Json.Linq.JProperty("source", Newtonsoft.Json.Linq.JValue.CreateNull()),
                    New Newtonsoft.Json.Linq.JProperty("host_operation", Newtonsoft.Json.Linq.JValue.CreateNull()),
                    New Newtonsoft.Json.Linq.JProperty("limit", Newtonsoft.Json.Linq.JValue.CreateNull()),
                    New Newtonsoft.Json.Linq.JProperty("observed", Newtonsoft.Json.Linq.JValue.CreateNull()),
                    New Newtonsoft.Json.Linq.JProperty("stack", New Newtonsoft.Json.Linq.JArray())))).ToString(Newtonsoft.Json.Formatting.None)
        End Function

        Private Shared Function GetConfiguredOptions() As PythonExecuteToolCoreOptions
            Dim options As PythonExecuteToolCoreOptions = GetConfiguredOptionsOrNothing()
            If options Is Nothing Then
                Throw New RedInkPythonAgentConfigurationException("PythonExecuteTool.Configure(...) must be called before registration or execution.")
            End If
            Return options
        End Function

        Private Shared Function GetConfiguredOptionsOrNothing() As PythonExecuteToolCoreOptions
            SyncLock ConfigurationLock
                Return ConfiguredOptions
            End SyncLock
        End Function
    End Class

End Namespace
