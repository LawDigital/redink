' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
Option Strict On
Option Explicit On
Option Infer On

Imports SharedLibrary.SharedLibrary.SharedContext

Namespace SharedLibrary
    Partial Public Class SharedMethods

        Public NotInheritable Class SemanticSearchRequestBudgetException
            Inherits System.InvalidOperationException
            Public Sub New(message As System.String)
                MyBase.New(message)
            End Sub
        End Class

        ' Marker used by the generic model helpers, including nested reader/OCR calls. They
        ' may update this private call state but must never write module restoration state.
        Friend Class IsolatedModelCallContext
            Inherits SharedContext
            Friend Property PinnedResolution As IsolatedSpecialTaskModel
            Friend Property RequestBudgetTokens As System.Int32
            Friend Property ReservedResponseTokens As System.Int32
        End Class

        Public NotInheritable Class IsolatedSpecialTaskModel
            Public ReadOnly Property Context As ISharedContext
            Public ReadOnly Property TaskName As System.String
            Public ReadOnly Property ModelName As System.String
            Public ReadOnly Property Signature As System.String
            Public ReadOnly Property UsedPrimaryFallback As System.Boolean
            Public ReadOnly Property UsesSecondApi As System.Boolean
            Public ReadOnly Property MaximumOutputTokens As System.Int32
            Public ReadOnly Property ContextWindowTokens As System.Int32
            Public ReadOnly Property TimeoutMilliseconds As System.Int64

            Friend Sub New(callContext As IsolatedModelCallContext, task As System.String,
                           secondApi As System.Boolean, contextTokens As System.Int32,
                           signatureValue As System.String)
                Context = callContext
                TaskName = task
                UsesSecondApi = secondApi
                UsedPrimaryFallback = Not secondApi
                ModelName = If(secondApi, callContext.INI_Model_2, callContext.INI_Model)
                MaximumOutputTokens = If(secondApi, callContext.INI_MaxOutputToken_2, callContext.INI_MaxOutputToken)
                ContextWindowTokens = contextTokens
                TimeoutMilliseconds = If(secondApi AndAlso callContext.INI_Timeout_2 > 0,
                                         callContext.INI_Timeout_2, callContext.INI_Timeout)
                Signature = signatureValue
                callContext.PinnedResolution = Me
            End Sub
        End Class

        ''' <summary>
        ''' Copies all interface properties, including mutable lists, without writing the host.
        ''' A repeated stable read detects concurrent configuration changes; it never restores
        ''' a snapshot over a newer chat configuration. Unknown future mutable types fail closed.
        ''' </summary>
        Public Shared Function CreateIsolatedModelCallContext(context As ISharedContext) As ISharedContext
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            Dim properties As System.Reflection.PropertyInfo() = GetType(ISharedContext).GetProperties()
            For attempt As System.Int32 = 1 To 3
                Dim snapshot As New IsolatedModelCallContext()
                For Each propertyInfo As System.Reflection.PropertyInfo In properties
                    Dim value As System.Object = propertyInfo.GetValue(context, Nothing)
                    If TypeOf value Is System.Collections.Generic.List(Of System.String) Then
                        value = New System.Collections.Generic.List(Of System.String)(DirectCast(value, System.Collections.Generic.List(Of System.String)))
                    ElseIf value IsNot Nothing AndAlso Not propertyInfo.PropertyType.IsValueType AndAlso propertyInfo.PropertyType IsNot GetType(System.String) Then
                        Throw New System.InvalidOperationException("A shared-context property requires an explicit isolated-copy contract: " & propertyInfo.Name)
                    End If
                    propertyInfo.SetValue(snapshot, value, Nothing)
                Next
                Dim stable As System.Boolean = True
                For Each propertyInfo As System.Reflection.PropertyInfo In properties
                    Dim before As System.Object = propertyInfo.GetValue(snapshot, Nothing)
                    Dim after As System.Object = propertyInfo.GetValue(context, Nothing)
                    If TypeOf before Is System.Collections.Generic.List(Of System.String) AndAlso TypeOf after Is System.Collections.Generic.List(Of System.String) Then
                        stable = System.Linq.Enumerable.SequenceEqual(
                            DirectCast(before, System.Collections.Generic.List(Of System.String)),
                            DirectCast(after, System.Collections.Generic.List(Of System.String)), System.StringComparer.Ordinal)
                    Else
                        stable = System.Object.Equals(before, after)
                    End If
                    If Not stable Then Exit For
                Next
                If stable Then
                    Dim previous As IsolatedModelCallContext = TryCast(context, IsolatedModelCallContext)
                    If previous IsNot Nothing Then snapshot.PinnedResolution = previous.PinnedResolution
                    Return snapshot
                End If
            Next
            Throw New System.InvalidOperationException("Model configuration changed while creating an isolated request; retry with a stable configuration.")
        End Function

        ''' <summary>
        ''' Resolves an accessible task assignment without mutating host or module state.
        ''' Primary is used only if no accessible assignment exists; read, parse, key-resolution,
        ''' invalid-assignment and later runtime failures are visible errors, never fallbacks.
        ''' The returned Context pins this resolution for a complete multi-call operation.
        ''' </summary>
        Public Shared Function ResolveIsolatedSpecialTaskModel(context As ISharedContext, specialTaskName As System.String) As IsolatedSpecialTaskModel
            If context Is Nothing Then Throw New System.ArgumentNullException(NameOf(context))
            If System.String.IsNullOrWhiteSpace(specialTaskName) Then Throw New System.ArgumentException("A special-task name is required.", NameOf(specialTaskName))
            Dim task As System.String = specialTaskName.Trim()
            Dim previous As IsolatedModelCallContext = TryCast(context, IsolatedModelCallContext)
            If previous IsNot Nothing AndAlso previous.PinnedResolution IsNot Nothing AndAlso
               System.String.Equals(previous.PinnedResolution.TaskName, task, System.StringComparison.OrdinalIgnoreCase) Then
                Return previous.PinnedResolution
            End If
            Dim isolated As IsolatedModelCallContext = DirectCast(CreateIsolatedModelCallContext(context), IsolatedModelCallContext)
            isolated.PinnedResolution = Nothing
            Dim model As ModelConfig = Nothing
            Dim contextWindow As System.Int32 = 0
            Dim found As System.Boolean = TryResolveIsolatedTaskAssignment(isolated, isolated.INI_AlternateModelPath, task, model, contextWindow)
            If found Then ApplyModelConfig(isolated, model)
            ValidateIsolatedModelConfiguration(isolated, found)
            Dim canonical As System.String = System.String.Join(Microsoft.VisualBasic.vbLf, New System.String() {
                "isolated-model-v1", task, found.ToString(System.Globalization.CultureInfo.InvariantCulture),
                If(found, isolated.INI_Model_2, isolated.INI_Model),
                If(found, isolated.INI_Endpoint_2, isolated.INI_Endpoint),
                If(found, isolated.INI_HeaderA_2, isolated.INI_HeaderA),
                If(found, isolated.INI_HeaderB_2, isolated.INI_HeaderB),
                If(found, isolated.INI_APICall_2, isolated.INI_APICall),
                If(found, isolated.INI_APICall_Object_2, isolated.INI_APICall_Object),
                If(found, isolated.INI_Response_2, isolated.INI_Response),
                If(found, isolated.INI_Temperature_2, isolated.INI_Temperature),
                If(found, isolated.INI_MaxOutputToken_2, isolated.INI_MaxOutputToken).ToString(System.Globalization.CultureInfo.InvariantCulture),
                contextWindow.ToString(System.Globalization.CultureInfo.InvariantCulture),
                isolated.INI_Model_Parameter1, isolated.INI_Model_Parameter2,
                isolated.INI_Model_Parameter3, isolated.INI_Model_Parameter4
            })
            Dim result As New IsolatedSpecialTaskModel(isolated, task, found, contextWindow,
                Global.SharedLibrary.Agents.TextExtractionResourceRegistry.HashString(canonical))
            System.Diagnostics.Debug.WriteLine("Isolated special-task resolution: task=" & task & "; model=" & result.ModelName &
                "; primary_fallback=" & result.UsedPrimaryFallback.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                "; signature=" & result.Signature)
            Return result
        End Function

        Private Shared Function TryResolveIsolatedTaskAssignment(
            context As ISharedContext, iniPath As System.String, task As System.String,
            ByRef model As ModelConfig, ByRef contextWindow As System.Int32
        ) As System.Boolean
            model = Nothing
            contextWindow = 0
            If System.String.IsNullOrWhiteSpace(iniPath) Then Return False
            Dim cacheHit As System.Boolean
            Dim sections As System.Collections.Generic.List(Of AlternativeModelIniSection) =
                GetAlternativeModelIniSections(iniPath, cacheHit, requireReadable:=True)
            For Each section As AlternativeModelIniSection In sections
                If section Is Nothing OrElse section.Values Is Nothing Then Continue For
                Dim raw As System.String = Nothing
                If Not section.Values.TryGetValue(task, raw) Then Continue For
                If Not IsModelAccessibleForCurrentUser(section.Values, context) Then Continue For
                If Not IsTruthyIniValue(raw) Then
                    Dim flag As System.String = If(raw, System.String.Empty).Split(";"c, "#"c)(0).Trim().Trim(System.Convert.ToChar(34), "'"c).ToLowerInvariant()
                    If flag.Length > 0 AndAlso flag <> "false" AndAlso flag <> "no" AndAlso flag <> "falsch" AndAlso
                       flag <> "nein" AndAlso flag <> "off" AndAlso flag <> "0" Then
                        Throw New System.InvalidOperationException("The special-task assignment flag is invalid: " & task)
                    End If
                    Continue For
                End If
                ValidateIsolatedNumericSetting(section.Values, "Timeout", allowZero:=True)
                ValidateIsolatedNumericSetting(section.Values, "MaxOutputToken", allowZero:=True)
                ValidateIsolatedNumericSetting(section.Values, "ContextWindowTokens", allowZero:=False)
                contextWindow = GetConfigInt(section.Values, "ContextWindowTokens", 0)
                model = CreateModelConfigFromDict(section.Values, context, section.Description, strictErrors:=True)
                If model Is Nothing Then Throw New System.InvalidOperationException("The configured task model could not be resolved: " & task)
                Return True
            Next
            Return False
        End Function

        Private Shared Sub ValidateIsolatedNumericSetting(values As System.Collections.Generic.Dictionary(Of System.String, System.String), key As System.String, allowZero As System.Boolean)
            Dim raw As System.String = GetConfigString(values, key)
            If System.String.IsNullOrWhiteSpace(raw) Then Return
            Dim parsed As System.Int32
            If Not System.Int32.TryParse(raw, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, parsed) OrElse
               parsed < If(allowZero, 0, 1) Then
                Throw New System.InvalidOperationException("The configured model has an invalid " & key & " value.")
            End If
        End Sub

        Private Shared Sub ValidateIsolatedModelConfiguration(context As ISharedContext, secondApi As System.Boolean)
            Dim endpoint As System.String = If(secondApi, context.INI_Endpoint_2, context.INI_Endpoint)
            Dim body As System.String = If(secondApi, context.INI_APICall_2, context.INI_APICall)
            Dim response As System.String = If(secondApi, context.INI_Response_2, context.INI_Response)
            Dim key As System.String = If(secondApi, context.DecodedAPI_2, context.DecodedAPI)
            If System.String.IsNullOrWhiteSpace(endpoint) OrElse System.String.IsNullOrWhiteSpace(body) OrElse System.String.IsNullOrWhiteSpace(response) Then
                Throw New System.InvalidOperationException("The resolved model configuration requires Endpoint, APICall and Response.")
            End If
            If System.Text.RegularExpressions.Regex.IsMatch(body & If(secondApi, context.INI_Model_2, context.INI_Model), "\{\s*parameter\d+\s*=", System.Text.RegularExpressions.RegexOptions.IgnoreCase) Then
                Throw New System.InvalidOperationException("The configured special-task model requires interactive parameter selection; supply fixed model settings for an isolated call.")
            End If
            Dim prefix As System.String = If(secondApi, context.INI_APIKeyPrefix_2, context.INI_APIKeyPrefix)
            Dim keyStatus As System.String = If(key, System.String.Empty)
            If Not System.String.IsNullOrEmpty(prefix) AndAlso keyStatus.StartsWith(prefix, System.StringComparison.Ordinal) Then
                keyStatus = keyStatus.Substring(prefix.Length).TrimStart()
            End If
            If Not System.String.IsNullOrWhiteSpace(keyStatus) AndAlso keyStatus.StartsWith("Error", System.StringComparison.OrdinalIgnoreCase) Then
                Throw New System.InvalidOperationException("The configured model credentials could not be resolved.")
            End If
            If If(secondApi, context.INI_APIEncrypted_2, context.INI_APIEncrypted) AndAlso
               Not If(secondApi, context.INI_OAuth2_2, context.INI_OAuth2) AndAlso System.String.IsNullOrWhiteSpace(key) Then
                Throw New System.InvalidOperationException("The configured encrypted model credentials could not be resolved.")
            End If
        End Sub

        ''' <summary>
        ''' Provider-neutral conservative budget: UTF-8 input bytes upper-bound byte-level
        ''' tokenizer input; include serialized prompts/templates, framing and output reservation.
        ''' The caller's configured request cap and a model's ContextWindowTokens both apply.
        ''' No question, card or source text is truncated to make a request fit.
        ''' </summary>
        Public Shared Function GetSemanticSearchRequestTokenBound(model As IsolatedSpecialTaskModel,
                                                                 systemPrompt As System.String,
                                                                 userPrompt As System.String,
                                                                 Optional reservedOutputTokens As System.Int32 = 4096) As System.Int64
            If model Is Nothing Then Throw New System.ArgumentNullException(NameOf(model))
            Dim template As System.String = If(model.UsesSecondApi, model.Context.INI_APICall_2, model.Context.INI_APICall)
            Dim endpoint As System.String = If(model.UsesSecondApi, model.Context.INI_Endpoint_2, model.Context.INI_Endpoint)
            Dim outputReserve As System.Int64 = System.Math.Max(CLng(System.Math.Max(0, model.MaximumOutputTokens)), CLng(System.Math.Max(0, reservedOutputTokens)))
            Dim requestTemplate As System.String = If(template, System.String.Empty) & If(endpoint, System.String.Empty)
            Dim systemCopies As System.Int32 = System.Math.Max(1, System.Text.RegularExpressions.Regex.Matches(requestTemplate, "\{promptsystem\}").Count)
            Dim userCopies As System.Int32 = System.Math.Max(1, System.Text.RegularExpressions.Regex.Matches(requestTemplate, "\{promptuser\}").Count)
            Return CLng(System.Text.Encoding.UTF8.GetByteCount(Newtonsoft.Json.JsonConvert.SerializeObject(If(systemPrompt, System.String.Empty)))) * systemCopies +
                CLng(System.Text.Encoding.UTF8.GetByteCount(Newtonsoft.Json.JsonConvert.SerializeObject(If(userPrompt, System.String.Empty)))) * userCopies +
                CLng(System.Text.Encoding.UTF8.GetByteCount(requestTemplate)) + outputReserve + 2048L
        End Function

        Friend Shared Sub ValidateIsolatedEndpointPromptBudget(context As ISharedContext, endpointTemplate As System.String,
                                                               systemPrompt As System.String, userPrompt As System.String)
            Dim isolated As IsolatedModelCallContext = TryCast(context, IsolatedModelCallContext)
            If isolated Is Nothing OrElse isolated.RequestBudgetTokens <= 0 Then Return
            If (If(endpointTemplate, System.String.Empty).Contains("{promptsystem}") AndAlso If(systemPrompt, System.String.Empty).Length > 32000) OrElse
               (If(endpointTemplate, System.String.Empty).Contains("{promptuser}") AndAlso If(userPrompt, System.String.Empty).Length > 32000) Then
                Throw New SemanticSearchRequestBudgetException("semantic_request_oversized: An endpoint prompt exceeds its existing 32000-character transport limit. The complete request was rejected before truncation.")
            End If
        End Sub

        Friend Shared Sub ValidateIsolatedSerializedRequestBudget(context As ISharedContext, endpoint As System.String, requestBody As System.String)
            Dim isolated As IsolatedModelCallContext = TryCast(context, IsolatedModelCallContext)
            If isolated Is Nothing OrElse isolated.RequestBudgetTokens <= 0 Then Return
            Dim bytes As System.Int64 = CLng(System.Text.Encoding.UTF8.GetByteCount(If(endpoint, System.String.Empty))) +
                CLng(System.Text.Encoding.UTF8.GetByteCount(If(requestBody, System.String.Empty))) +
                CLng(isolated.ReservedResponseTokens) + 2048L
            If bytes > isolated.RequestBudgetTokens Then
                Throw New SemanticSearchRequestBudgetException("semantic_request_oversized: The fully expanded request and reserved response exceed the configured context budget.")
            End If
        End Sub

        Public Shared Sub ValidateSemanticSearchRequestBudget(model As IsolatedSpecialTaskModel,
                                                              systemPrompt As System.String, userPrompt As System.String,
                                                              maximumRequestTokens As System.Int32,
                                                              Optional reservedOutputTokens As System.Int32 = 4096)
            If maximumRequestTokens < 1 Then Throw New System.ArgumentOutOfRangeException(NameOf(maximumRequestTokens))
            Dim limit As System.Int64 = maximumRequestTokens
            If model.ContextWindowTokens > 0 Then limit = System.Math.Min(limit, model.ContextWindowTokens)
            Dim bound As System.Int64 = GetSemanticSearchRequestTokenBound(model, systemPrompt, userPrompt, reservedOutputTokens)
            If bound > limit Then
                Throw New SemanticSearchRequestBudgetException("semantic_request_oversized: Complete request bound " &
                    bound.ToString(System.Globalization.CultureInfo.InvariantCulture) & " exceeds the configured context budget " &
                    limit.ToString(System.Globalization.CultureInfo.InvariantCulture) & "; split the authorized records or source at a host-owned boundary.")
            End If
        End Sub
    End Class
End Namespace
