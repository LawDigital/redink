' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. For license to use see https://redink.ai.
' Reuses the existing primary/secondary/alternate model configuration contract.
Option Explicit On
Option Strict On

Namespace Global.SharedLibrary.TabularReview
    Public NotInheritable Class ReviewModelChoice
        Public Property Key As System.String
        Public Property Label As System.String
        <Newtonsoft.Json.JsonIgnore>
        Public Property PrimaryNumber As System.Int32
        <Newtonsoft.Json.JsonIgnore>
        Public Property Secondary As System.Boolean
        <Newtonsoft.Json.JsonIgnore>
        Public Property Configuration As Global.SharedLibrary.SharedLibrary.ModelConfig
    End Class

    Public NotInheritable Class ReviewModelCatalog
        Private Sub New()
        End Sub

        Public Shared Function Load(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                                    Optional log As System.Action(Of System.String) = Nothing) As System.Collections.Generic.List(Of ReviewModelChoice)
            Dim result As New System.Collections.Generic.List(Of ReviewModelChoice)()
            For Each number As System.Int32 In Global.SharedLibrary.SharedLibrary.PrimaryModelManager.GetAvailableModels()
                result.Add(New ReviewModelChoice With {
                    .Key = "primary:" & number.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    .Label = Global.SharedLibrary.SharedLibrary.PrimaryModelManager.GetModelDisplayName(number), .PrimaryNumber = number
                })
            Next
            If result.Count = 0 Then result.Add(New ReviewModelChoice With {.Key = "current", .Label = "Current model — " & context.INI_Model})
            If context.INI_SecondAPI AndAlso Not System.String.IsNullOrWhiteSpace(context.INI_Model_2) Then
                result.Add(New ReviewModelChoice With {.Key = "secondary", .Label = "Secondary — " & context.INI_Model_2, .Secondary = True})
            End If
            If Not System.String.IsNullOrWhiteSpace(context.INI_AlternateModelPath) Then
                Try
                    Dim isolated As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext = Global.SharedLibrary.SharedLibrary.SharedMethods.CreateIsolatedModelCallContext(context)
                    Dim path As System.String = Global.SharedLibrary.SharedLibrary.SharedMethods.ExpandEnvironmentVariables(context.INI_AlternateModelPath)
                    For Each config As Global.SharedLibrary.SharedLibrary.ModelConfig In Global.SharedLibrary.SharedLibrary.SharedMethods.LoadAlternativeModels(path, isolated)
                        Dim identity As System.String = System.String.Join(Microsoft.VisualBasic.Constants.vbLf, New System.String() {
                            If(config.ModelDescription, ""), If(config.Model, ""), If(config.Endpoint, ""), If(config.APICall, ""), If(config.Response, "")
                        })
                        Dim key As System.String = "alternate:" & ReviewContract.Digest(identity)
                        If result.Exists(Function(choice As ReviewModelChoice) choice.Key = key) Then Continue For
                        result.Add(New ReviewModelChoice With {
                            .Key = key, .Label = "Alternate — " & If(System.String.IsNullOrWhiteSpace(config.ModelDescription), config.Model, config.ModelDescription),
                            .Secondary = True, .Configuration = config
                        })
                    Next
                Catch ex As System.Exception
                    If log IsNot Nothing Then log("Alternate models could not be loaded: " & ex.Message)
                End Try
            End If
            Return result
        End Function

        Public Shared Function CreateInvoker(context As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext,
                                             choice As ReviewModelChoice) As System.Func(Of System.String, System.String, System.Threading.CancellationToken, System.Threading.Tasks.Task(Of System.String))
            If choice Is Nothing Then Throw New System.InvalidOperationException("Select an available model before generating an answer.")
            Dim isolated As Global.SharedLibrary.SharedLibrary.SharedContext.ISharedContext
            If choice.PrimaryNumber > 0 Then
                isolated = Global.SharedLibrary.SharedLibrary.PrimaryModelManager.CreateIsolatedPrimaryModelContext(context, choice.PrimaryNumber)
            Else
                isolated = Global.SharedLibrary.SharedLibrary.SharedMethods.CreateIsolatedModelCallContext(context)
            End If
            If choice.Configuration IsNot Nothing Then
                Dim invalid As System.Boolean = False
                Global.SharedLibrary.SharedLibrary.SharedMethods.ApplyModelConfig(isolated, choice.Configuration, invalid)
                If invalid Then Throw New System.InvalidOperationException("The selected alternate model configuration is invalid.")
            End If
            Dim secondary As System.Boolean = choice.Secondary
            Return Function(systemPrompt As System.String, userPrompt As System.String, cancellation As System.Threading.CancellationToken)
                       Return Global.SharedLibrary.SharedLibrary.SharedMethods.LLM(isolated, systemPrompt, userPrompt,
                           Timeout:=If(secondary AndAlso isolated.INI_Timeout_2 > 0, isolated.INI_Timeout_2, isolated.INI_Timeout),
                           UseSecondAPI:=secondary, Hidesplash:=True, cancellationToken:=cancellation)
                   End Function
        End Function
    End Class
End Namespace
