' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: RedinkPythonAgentHostService.vb
' Purpose: Defines safe host-service dispatch for `python_execute`, including
'          typed host-call failures and delegate-based capability routing.
'
' Architecture / How it works:
'  - `RedInkPythonAgentHostCallException` provides typed, model-safe host error
'    signaling so raw host exceptions are not exposed back to the model.
'  - `RedInkPythonAgentDelegatingHostServiceHandler` is a capability gate that
'    only exposes operations whose delegates are explicitly wired by the host.
'  - `HandleAsync()` dispatches by `RedInkPythonAgentHostOperation`, validates
'    argument payload types, invokes the matching delegate, and converts runtime
'    failures into deterministic safe response envelopes.
'  - Unsupported or disabled operations always fail with
'    `HOST_OPERATION_NOT_ALLOWED`, keeping sandbox capabilities aligned with the
'    active tooling loop.
' =============================================================================
Option Explicit On
Option Infer On

Imports System.Threading
Imports System.Threading.Tasks

Namespace Agents

    ''' <summary>
    ''' Raised by a host-service delegate to return a typed, model-safe failure code
    ''' (for example WEB_CONTENT_TYPE_UNSUPPORTED) instead of leaking raw exception text.
    ''' </summary>
    Public NotInheritable Class RedInkPythonAgentHostCallException
        Inherits System.Exception
        Public ReadOnly Property Code As System.String
        Public ReadOnly Property Retryable As System.Boolean
        Public Sub New(code As System.String, Optional retryable As System.Boolean = False, Optional message As System.String = Nothing)
            MyBase.New(If(message, code))
            Me.Code = code
            Me.Retryable = retryable
        End Sub
    End Class

    ''' <summary>
    ''' Host-agnostic host-service handler for python_execute. Each capability is wired only
    ''' when the surrounding tooling loop already exposes the corresponding tool, so the sandbox
    ''' can never reach a capability the loop itself does not have. LLM is always wired; web.get
    ''' and web.search are wired only when their delegates are provided. Any operation whose
    ''' delegate is absent is rejected deterministically with HOST_OPERATION_NOT_ALLOWED, so a
    ''' disabled call fails even if the worker attempts it.
    ''' </summary>
    Public NotInheritable Class RedInkPythonAgentDelegatingHostServiceHandler
        Implements IRedInkPythonAgentHostServiceHandler

        ''' <summary>Always set: routes llm.complete to the host LLM. (system, user, ct) -> text.</summary>
        Public Property LlmAsync As System.Func(Of System.String, System.String, CancellationToken, Task(Of System.String))

        ''' <summary>Set only when web retrieval is on for the loop: (url, maxChars, ct) -> extracted text.</summary>
        Public Property WebGetAsync As System.Func(Of System.String, System.Int32, CancellationToken, Task(Of System.String))

        ''' <summary>Set only when web search is on for the loop: (query, maxResults, ct) -> results.</summary>
        Public Property WebSearchAsync As System.Func(Of System.String, System.Int32, CancellationToken, Task(Of System.Collections.Generic.IReadOnlyList(Of RedInkPythonAgentWebSearchItem)))

        Public Async Function HandleAsync(request As RedInkPythonAgentHostCallRequest, cancellationToken As CancellationToken) As Task(Of RedInkPythonAgentHostCallResponse) Implements IRedInkPythonAgentHostServiceHandler.HandleAsync
            If request Is Nothing Then
                Return RedInkPythonAgentHostCallResponse.Failure("HOST_REQUEST_INVALID", False, Nothing)
            End If

            Dim failure As RedInkPythonAgentHostCallResponse = Nothing
            Try
                Select Case request.Operation

                    Case RedInkPythonAgentHostOperation.LlmComplete
                        If LlmAsync Is Nothing Then Return NotAllowed()
                        Dim args = TryCast(request.Arguments, RedInkPythonAgentLlmRequest)
                        If args Is Nothing Then Return RedInkPythonAgentHostCallResponse.Failure("HOST_REQUEST_INVALID", False, Nothing)
                        Dim text As System.String = Await LlmAsync(args.SystemPrompt, args.UserPrompt, cancellationToken).ConfigureAwait(False)
                        Return RedInkPythonAgentHostCallResponse.SuccessLlm(If(text, System.String.Empty))

                    Case RedInkPythonAgentHostOperation.WebGet
                        If WebGetAsync Is Nothing Then Return NotAllowed()
                        Dim args = TryCast(request.Arguments, RedInkPythonAgentWebGetRequest)
                        If args Is Nothing Then Return RedInkPythonAgentHostCallResponse.Failure("HOST_REQUEST_INVALID", False, Nothing)
                        Dim text As System.String = Await WebGetAsync(args.Url, args.MaximumCharacters, cancellationToken).ConfigureAwait(False)
                        Return RedInkPythonAgentHostCallResponse.SuccessWebGet(If(text, System.String.Empty))

                    Case RedInkPythonAgentHostOperation.WebSearch
                        If WebSearchAsync Is Nothing Then Return NotAllowed()
                        Dim args = TryCast(request.Arguments, RedInkPythonAgentWebSearchRequest)
                        If args Is Nothing Then Return RedInkPythonAgentHostCallResponse.Failure("HOST_REQUEST_INVALID", False, Nothing)
                        Dim results = Await WebSearchAsync(args.Query, args.MaximumResults, cancellationToken).ConfigureAwait(False)
                        Return RedInkPythonAgentHostCallResponse.SuccessWebSearch(If(results, New System.Collections.Generic.List(Of RedInkPythonAgentWebSearchItem)()))

                        ' ─────────────────────────────────────────────────────────────────────────
                        ' DORMANT: web.download is intentionally NOT wired yet.
                        ' It is deliberately absent from the RedInkPythonAgentHostOperation enum,
                        ' never added to AllowedOperations, and never advertised to the model. This
                        ' branch is scaffolding only, pending the python-agent (LPAC worker) side
                        ' implementing the client call. Expected future protocol-v2 worker call once
                        ' the worker exposes it:
                        '
                        '     saved = agent_api.web_download(
                        '         "https://example.org/report.pdf",
                        '         target_name="report.pdf")   ' writes the ORIGINAL binary bytes into
                        '                                       ' the agent output area via output_path(...)
                        '
                        ' When implemented, add RedInkPythonAgentHostOperation.WebDownload, a
                        ' WebDownloadAsync delegate here, and gate it on download_web_files being
                        ' selected for the loop. Until then this operation must remain unreachable.
                        ' ─────────────────────────────────────────────────────────────────────────

                    Case Else
                        Return NotAllowed()
                End Select

            Catch ex As OperationCanceledException
                Throw
            Catch ex As RedInkPythonAgentHostCallException
                failure = RedInkPythonAgentHostCallResponse.Failure(ex.Code, ex.Retryable, Nothing)
            Catch ex As System.Exception
                System.Diagnostics.Trace.WriteLine(ex.ToString())
                failure = RedInkPythonAgentHostCallResponse.Failure("HOST_CALL_FAILED", False, Nothing)
            End Try

            Return failure
        End Function

        Private Shared Function NotAllowed() As RedInkPythonAgentHostCallResponse
            Return RedInkPythonAgentHostCallResponse.Failure("HOST_OPERATION_NOT_ALLOWED", False, Nothing)
        End Function

    End Class

End Namespace
