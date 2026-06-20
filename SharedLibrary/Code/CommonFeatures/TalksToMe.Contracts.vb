' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: TalksToMe.Contracts.vb
' Purpose: Defines the contracts for the "Talk to Me" feature, ensuring a clean
'          separation between the UI, the host application, and the transcription
'          backend.
'
' Architecture:
'  - ITalkToMeHost: An interface that must be implemented by the host
'    application (e.g., Word Add-in). It defines callbacks for UI events,
'    such as toggling recording.
'  - ITranscriptionEngine: An interface for transcription services, abstracting
'    the specific implementation (e.g., Google, Azure) from the rest of the
'    application. It defines methods for starting and stopping transcription
'    and events for results and errors.
'  - TranscriptionResult: A class that encapsulates the data returned from a
'    transcription engine, including the transcribed text and a confidence score.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Threading
Imports System.Threading.Tasks

Namespace SharedLibrary
    Public Enum TalkToMeActionKind
        None
        HostCommand
        TypeText
        InsertText
        Freestyle
        GotoText
        FindText
    End Enum

    Public Class TalkToMeCommandDefinition
        Public Property Name As String = ""
        Public Property Label As String = ""
        Public Property Category As String = ""
        Public Property Description As String = ""
        Public Property Aliases As List(Of String)

        Public Sub New()
            Aliases = New List(Of String)()
        End Sub
    End Class

    Public Class TalkToMeDocumentContext
        Public Property HostName As String = ""
        Public Property DocumentName As String = ""
        Public Property DocumentText As String = ""
        Public Property SelectionText As String = ""
        Public Property CursorContext As String = ""
        Public Property CaretPosition As String = ""
        Public Property HasSelection As Boolean
        Public Property ActiveSurface As String = ""
        Public Property CanWriteToDocument As Boolean
    End Class

    Public Class TalkToMeStructuredResponse
        Public Property Action As String = ""
        Public Property HostCommandName As String = ""
        Public Property Text As String = ""
        Public Property Query As String = ""
        Public Property Instruction As String = ""
        Public Property Reason As String = ""
    End Class

    Public Class TalkToMeDispatchResult
        Public Property Handled As Boolean
        Public Property StatusText As String = ""
        Public Property TranscriptToDisplay As String = ""
    End Class

    Public Class TalkToMeSpeechConfigurationResult
        Public Property Applied As Boolean
        Public Property IncludeFullDocument As Boolean
        Public Property Summary As String = ""
    End Class

    Public Class TalkToMeTranscriptEventArgs
        Inherits EventArgs

        Public ReadOnly Property Text As String

        Public Sub New(text As String)
            Me.Text = If(text, "")
        End Sub
    End Class

    Public Interface ITalkToMeHostAdapter
        ReadOnly Property HostName As String
        Function GetSupportedCommands() As List(Of TalkToMeCommandDefinition)
        Function GetPromptContext(includeFullDocument As Boolean) As TalkToMeDocumentContext
        Function ResolveWithLlmAsync(spokenInstruction As String,
                                     context As TalkToMeDocumentContext,
                                     supportedCommands As List(Of TalkToMeCommandDefinition),
                                     cancellationToken As CancellationToken) As Task(Of TalkToMeStructuredResponse)
        Function ExecuteAsync(response As TalkToMeStructuredResponse,
                              cancellationToken As CancellationToken) As Task(Of TalkToMeDispatchResult)
    End Interface

    Public Interface ITalkToMeSpeechAdapter
        Event PartialTranscriptReceived As EventHandler(Of TalkToMeTranscriptEventArgs)
        Event FinalTranscriptReceived As EventHandler(Of TalkToMeTranscriptEventArgs)

        ReadOnly Property IsListening As Boolean
        ReadOnly Property IsConfigured As Boolean
        ReadOnly Property IsSpeechOutputAvailable As Boolean
        ReadOnly Property IsSpeechOutputEnabled As Boolean
        ReadOnly Property IsSpeechOutputActive As Boolean
        ReadOnly Property CanAcceptExternalSpeech As Boolean

        Function Configure(owner As System.Windows.Forms.IWin32Window,
                           currentIncludeFullDocument As Boolean) As TalkToMeSpeechConfigurationResult
        Function StartListeningAsync(cancellationToken As CancellationToken) As Task
        Function StopListeningAsync() As Task
        Function GetConfigurationSummary() As String
        Function GetSpeechOutputSummary() As String
        Function ConfigureSpeechOutput(owner As System.Windows.Forms.IWin32Window) As String
        Function ToggleSpeechOutputEnabled() As Boolean
        Function SubmitExternalSpeechAsync(speakerName As String,
                                           text As String,
                                           cancellationToken As CancellationToken) As Task(Of Boolean)
    End Interface

    Public Class TalkToMeCoordinator
        Private ReadOnly _host As ITalkToMeHostAdapter
        Private ReadOnly _includeFullDocumentProvider As Func(Of Boolean)

        Public Sub New(host As ITalkToMeHostAdapter,
                       includeFullDocumentProvider As Func(Of Boolean))
            _host = host
            _includeFullDocumentProvider = includeFullDocumentProvider
        End Sub

        Public Async Function ProcessTranscriptAsync(spokenText As String,
                                                     cancellationToken As CancellationToken) As Task(Of TalkToMeDispatchResult)
            Dim rawText As String = If(spokenText, "").Trim()

            If String.IsNullOrWhiteSpace(rawText) Then
                Return New TalkToMeDispatchResult With {
                    .Handled = False,
                    .StatusText = "No speech detected.",
                    .TranscriptToDisplay = ""
                }
            End If

            Dim includeFullDocument As Boolean = False
            If _includeFullDocumentProvider IsNot Nothing Then
                includeFullDocument = _includeFullDocumentProvider.Invoke()
            End If

            Dim supportedCommands As List(Of TalkToMeCommandDefinition) = _host.GetSupportedCommands()
            Dim context As TalkToMeDocumentContext = _host.GetPromptContext(includeFullDocument)

            Dim llmResponse As TalkToMeStructuredResponse =
                Await _host.ResolveWithLlmAsync(rawText, context, supportedCommands, cancellationToken).ConfigureAwait(False)

            If llmResponse Is Nothing Then
                Return New TalkToMeDispatchResult With {
                    .Handled = False,
                    .StatusText = "No actionable response was returned.",
                    .TranscriptToDisplay = rawText
                }
            End If

            If String.IsNullOrWhiteSpace(llmResponse.Action) Then
                llmResponse.Action = "none"
            End If

            If String.IsNullOrWhiteSpace(llmResponse.Instruction) Then
                llmResponse.Instruction = rawText
            End If

            Return Await _host.ExecuteAsync(llmResponse, cancellationToken).ConfigureAwait(False)
        End Function

    End Class
End Namespace
