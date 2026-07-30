' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: IWordDocumentHost.vb
' Purpose: Bridge so SharedLibrary can address the currently-open Word document
'          without taking a direct dependency on Microsoft.Office.Interop.Word.
'          The Word host implements this; Outlook implements a stub.
'
' Operations (all SYNCHRONOUS by design):
'  - Read verbs: HasActiveDocument, ListOpenDocuments, GetActiveDocument,
'    ExtractText, SearchJson, ListComments (always available).
'  - Write verbs: InsertTextJson, ReplaceJson, DeleteJson, AddCommentJson, FormatJson
'    (gated by WordHostPolicy.ActiveDocReadOnly).
'  - Host enforces read-only policy; COM/STA calls do not benefit from async.
' =============================================================================

Option Strict On
Option Explicit On

Namespace Agents

    Public Class OpenDocInfo
        Public Property Name As String
        Public Property Path As String
        Public Property IsActive As Boolean
        Public Property IsReadOnly As Boolean
        Public Property IsSaved As Boolean
    End Class

    Public Class WordDocComment
        Public Property Id As String
        Public Property Author As String
        Public Property Initials As String
        Public Property Date_ As DateTime?
        Public Property Text As String
        Public Property AnchorText As String
    End Class

    Public Interface IWordDocumentHost
        ''' <summary>True if Word is reachable and at least one document is open.</summary>
        Function HasActiveDocument() As Boolean

        ''' <summary>Lists all open documents.</summary>
        Function ListOpenDocuments() As List(Of OpenDocInfo)

        ''' <summary>Metadata for the active document, or Nothing.</summary>
        Function GetActiveDocument() As OpenDocInfo

        ''' <summary>Returns the document's plain text (active or by exact name/path).</summary>
        Function ExtractText(documentNameOrPath As String, maxChars As Integer) As String

        ''' <summary>Search the active or named document for substring/regex. Returns serialized JSON.</summary>
        Function SearchJson(documentNameOrPath As String, query As String,
                            useRegex As Boolean, ignoreCase As Boolean, maxHits As Integer) As String

        ''' <summary>Lists comments on the named/active document.</summary>
        Function ListComments(documentNameOrPath As String) As List(Of WordDocComment)

        ' --- mutating operations (host enforces WordHostPolicy.ActiveDocReadOnly) ---

        ''' <summary>Inserts text at the document's current selection / end. Returns JSON result.</summary>
        Function InsertTextJson(documentNameOrPath As String, text As String, location As String,
                                trackChanges As Boolean) As String

        ''' <summary>Replaces the resilient-located 'find' span with 'replacement' (optionally scope-expanded).</summary>
        Function ReplaceJson(documentNameOrPath As String, find As String, replacement As String,
                              onlyFirst As Boolean, trackChanges As Boolean, matchScope As String) As String

        ''' <summary>Deletes the resilient-located 'find' span (or its enclosing sentence/paragraph).</summary>
        Function DeleteJson(documentNameOrPath As String, find As String,
                             trackChanges As Boolean, matchScope As String) As String

        ''' <summary>Adds a comment anchored to a 'find' match.</summary>
        Function AddCommentJson(documentNameOrPath As String, find As String, text As String,
                                 author As String, initials As String) As String

        ''' <summary>Applies paragraph/run formatting on a 'find' match.</summary>
        Function FormatJson(documentNameOrPath As String, find As String, styleId As String,
                             bold As Boolean?, italic As Boolean?, underline As Boolean?,
                             sizePt As Integer, color As String, align As String) As String
    End Interface

    Public NotInheritable Class WordHostPolicy

        Private Sub New()
        End Sub

        ''' <summary>
        ''' When True (default), only the read-only worddoc_* verbs are registered and the
        ''' mutating verbs refuse to run. Flip to False (via future ribbon toggle) to let
        ''' the agent modify the open document via Word interop.
        ''' </summary>
        Public Shared Property ActiveDocReadOnly As Boolean = False

        ''' <summary>
        ''' Host registration. Set once at startup; null for Outlook (no Word interop).
        ''' </summary>
        Public Shared Property Host As IWordDocumentHost = Nothing
    End Class

End Namespace
