Option Explicit On
Option Strict On
Option Infer On

Namespace SharedLibrary
    Partial Public Class SharedMethods
        Public Shared Function TranslationPreferencePrompt(context As SharedContext.ISharedContext, style As System.String, length As System.String) As System.String
            Dim parts As New System.Collections.Generic.List(Of System.String)()
            Select Case If(style, "").Trim().ToLowerInvariant()
                Case "", "default"
                Case "close"
                    parts.Add(context.SP_Translate_Style_Close)
                Case "natural"
                    parts.Add(context.SP_Translate_Style_Natural)
                Case "adapt"
                    parts.Add(context.SP_Translate_Style_Adapt)
                Case Else
                    Throw New System.ArgumentException("Unknown translation style: " & style)
            End Select
            Select Case If(length, "").Trim().ToLowerInvariant()
                Case "", "default"
                Case "flexible"
                    parts.Add(context.SP_Translate_Length_Flexible)
                Case "similar"
                    parts.Add(context.SP_Translate_Length_Similar)
                Case "fit"
                    parts.Add(context.SP_Translate_Length_Fit)
                Case Else
                    Throw New System.ArgumentException("Unknown translation length preference: " & length)
            End Select
            If parts.Count = 0 Then Return System.String.Empty
            Return System.Environment.NewLine & System.String.Join(System.Environment.NewLine, parts)
        End Function

        Public Shared Function ChooseTranslationOptions(context As SharedContext.ISharedContext, title As System.String, ByRef cancelled As System.Boolean, Optional ByRef checkFit As System.Boolean = False, Optional ByRef preferWording As System.Boolean = False) As System.String
            Dim styleLabels As System.String() = {"Use configured prompt", "Close to source", "Natural", "Adapt to audience"}
            Dim lengthLabels As System.String() = {"Use configured prompt", "Flexible", "Similar length", "Fit original space"}
            Dim styles As System.String() = {"default", "close", "natural", "adapt"}
            Dim lengths As System.String() = {"default", "flexible", "similar", "fit"}
            Dim parameters As InputParameter() = {
                New InputParameter("Translation style", styleLabels(0), styleLabels),
                New InputParameter("Length preference", lengthLabels(0), lengthLabels)
            }
            cancelled = Not ShowCustomVariableInputForm("Choose how to translate the text and handle its length.", title, parameters)
            If cancelled Then Return System.String.Empty

            Dim style = styles(System.Array.IndexOf(styleLabels, CStr(parameters(0).Value)))
            Dim length = lengths(System.Array.IndexOf(lengthLabels, CStr(parameters(1).Value)))
            preferWording = style = "close"
            checkFit = length = "fit"
            Return TranslationPreferencePrompt(context, style, length)
        End Function
    End Class
End Namespace
