' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365Models.vb
' Purpose: Shared types, enums and option classes used by M365Service.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Collections.Generic
Imports Newtonsoft.Json.Linq

Namespace SharedLibrary

    ''' <summary>
    ''' Sources that can be searched against Microsoft 365 / Graph.
    ''' Combine with bitwise Or, e.g. <c>M365SearchSources.Mail Or M365SearchSources.SharePoint</c>.
    ''' </summary>
    <Flags>
    Public Enum M365SearchSources
        None = 0
        Mail = 1
        OneDrive = 2          ' personal driveItem hits
        SharePoint = 4        ' SP driveItem (document libraries)
        SharePointSites = 8
        SharePointListItems = 16
        Teams = 32            ' chatMessage
        Calendar = 64         ' event
        OneNote = 128         ' /me/onenote/pages?$search=
        People = 256          ' /me/people?$search=
        AllFiles = OneDrive Or SharePoint
        AllSharePoint = SharePoint Or SharePointSites Or SharePointListItems
        All = Mail Or OneDrive Or SharePoint Or SharePointSites Or SharePointListItems _
              Or Teams Or Calendar Or OneNote Or People
    End Enum

    ''' <summary>
    ''' Stages reported via <see cref="IProgress(Of M365SearchProgress)"/>.
    ''' </summary>
    Public Enum M365ProgressStage
        SigningIn
        Searching
        ParsingResults
        Completed
        Failed
        Cancelled
    End Enum

    ''' <summary>Progress payload reported during <see cref="M365Service.SearchAsync"/>.</summary>
    Public Structure M365SearchProgress
        Public Property Stage As M365ProgressStage
        Public Property Source As M365SearchSources
        Public Property Message As String
        ''' <summary>0..100, or -1 for indeterminate.</summary>
        Public Property Percent As Integer
        Public Property HitsSoFar As Integer
    End Structure

    ''' <summary>Single search hit normalised across sources.</summary>
    Public Class M365SearchHit
        Public Property Source As M365SearchSources
        ''' <summary>Graph entity id (messageId, driveItemId, eventId, chatMessage id, …).</summary>
        Public Property Id As String
        ''' <summary>Parent id for sources that need it (e.g. chatId for chatMessage, siteId for listItem, driveId for driveItem).</summary>
        Public Property ParentId As String
        Public Property Title As String
        ''' <summary>HTML/plain summary returned by Graph (already stripped of &lt;c0&gt; markers when possible).</summary>
        Public Property Summary As String
        Public Property WebUrl As String
        Public Property LastModifiedUtc As DateTime?
        Public Property Author As String          ' from / createdBy.user.displayName / organizer.emailAddress
        Public Property AdditionalText As String  ' source-specific extras (e.g. attendee count)
        ''' <summary>Original JSON for power users / debugging.</summary>
        Public Property RawJson As JObject
    End Class

    ''' <summary>Aggregate result returned by <see cref="M365Service.SearchAsync"/>.</summary>
    Public Class M365SearchResult
        Public Property Query As String
        Public Property RequestedSources As M365SearchSources
        Public Property Hits As New List(Of M365SearchHit)()
        ''' <summary>Per-source hit count (only sources actually queried).</summary>
        Public Property CountsBySource As New Dictionary(Of M365SearchSources, Integer)()
        ''' <summary>Per-source error message; absent key = no error.</summary>
        Public Property ErrorsBySource As New Dictionary(Of M365SearchSources, String)()
        Public Property TotalEstimated As Integer
        Public Property StartedUtc As DateTime
        Public Property FinishedUtc As DateTime
        Public Property NextFromIndexBySource As New System.Collections.Generic.Dictionary(Of M365SearchSources, System.Int32)()
        Public Property ExhaustedBySource As New System.Collections.Generic.Dictionary(Of M365SearchSources, System.Boolean)()

        Public ReadOnly Property HasErrors As Boolean
            Get
                Return ErrorsBySource IsNot Nothing AndAlso ErrorsBySource.Count > 0
            End Get
        End Property
    End Class

    ''' <summary>Options for <see cref="M365Service.SearchAsync"/>.</summary>
    Public Class M365SearchOptions
        ''' <summary>Maximum hits returned per source (up to 2500; default 25). Mail/event Graph pages are at most 25.</summary>
        Public Property MaxPerSource As Integer = 25
        ''' <summary>Optional zero-based offset for paged search retrieval.</summary>
        Public Property FromIndex As Integer = 0
        ''' <summary>Optional KQL date floor (uses Graph KQL <c>received&gt;=</c> for messages, <c>lastModifiedDateTime&gt;=</c> for files).</summary>
        Public Property [From] As Date?
        Public Property [To] As Date?
        ''' <summary>Optional extra KQL appended verbatim to the queryString (e.g. <c>from:alice@contoso.com</c>).</summary>
        Public Property KqlExtra As String
        ''' <summary>If True, Graph hit summaries are stripped of <c0> markers and HTML.</summary>
        Public Property CleanSummaries As Boolean = True
        ''' <summary>If True, sources are queried in parallel (default). Set False for stricter rate-limit behaviour.</summary>
        Public Property Parallel As Boolean = True
    End Class

    ''' <summary>Fields to include when retrieving an Outlook message.</summary>
    <Flags>
    Public Enum M365MessageFields
        Headers = 0           ' subject, from, sent, importance — always returned
        Body = 1
        Recipients = 2        ' toRecipients/cc/bcc
        AttachmentsList = 4   ' metadata only
        InternetHeaders = 8
        Categories = 16
        All = Body Or Recipients Or AttachmentsList Or InternetHeaders Or Categories
    End Enum

    ''' <summary>Full Outlook message DTO returned by <see cref="M365Service.GetMessageAsync"/>.</summary>
    Public Class M365Message
        Public Property Id As String
        Public Property Subject As String
        Public Property From As String
        Public Property FromAddress As String
        Public Property ReceivedUtc As DateTime?
        Public Property SentUtc As DateTime?
        Public Property Importance As String
        Public Property HasAttachments As Boolean
        Public Property BodyContentType As String  ' "html" or "text"
        Public Property Body As String
        Public Property BodyPreview As String
        Public Property To_ As New List(Of String)()
        Public Property Cc As New List(Of String)()
        Public Property Bcc As New List(Of String)()
        Public Property Categories As New List(Of String)()
        Public Property InternetMessageId As String
        Public Property ConversationId As String
        Public Property WebLink As String
        Public Property Attachments As New List(Of M365AttachmentInfo)()
        Public Property RawJson As JObject
    End Class

    Public Class M365AttachmentInfo
        Public Property Id As String
        Public Property Name As String
        Public Property ContentType As String
        Public Property Size As Long
        Public Property IsInline As Boolean
        Public Property OdataType As String   ' fileAttachment / itemAttachment / referenceAttachment
    End Class

    ''' <summary>DriveItem (file) DTO.</summary>
    Public Class M365DriveItem
        Public Property Id As String
        Public Property DriveId As String
        Public Property Name As String
        Public Property Path As String
        Public Property MimeType As String
        Public Property Size As Long
        Public Property IsFolder As Boolean
        Public Property WebUrl As String
        Public Property CreatedBy As String
        Public Property LastModifiedBy As String
        Public Property CreatedUtc As DateTime?
        Public Property LastModifiedUtc As DateTime?
        Public Property ETag As String
        Public Property RawJson As JObject
    End Class

    Public Class M365Event
        Public Property Id As String
        Public Property Subject As String
        Public Property Organizer As String
        Public Property Location As String
        Public Property StartUtc As DateTime?
        Public Property EndUtc As DateTime?
        Public Property IsAllDay As Boolean
        Public Property BodyPreview As String
        Public Property WebLink As String
        Public Property Attendees As New List(Of String)()
        Public Property RawJson As JObject
    End Class

    Public Class M365ChatMessage
        Public Property Id As String
        Public Property ChatId As String
        Public Property TeamId As String
        Public Property ChannelId As String
        Public Property From As String
        Public Property CreatedUtc As DateTime?
        Public Property ContentType As String
        Public Property Content As String
        Public Property WebUrl As String
        Public Property RawJson As JObject
    End Class

    Public Class M365OneNotePage
        Public Property Id As String
        Public Property Title As String
        Public Property CreatedUtc As DateTime?
        Public Property LastModifiedUtc As DateTime?
        Public Property ContentUrl As String
        Public Property HtmlContent As String   ' populated by GetOneNotePageAsync(includeContent:=True)
        Public Property RawJson As JObject
    End Class

    Public Class M365Person
        Public Property Id As String
        Public Property DisplayName As String
        Public Property JobTitle As String
        Public Property CompanyName As String
        Public Property EmailAddresses As New List(Of String)()
        Public Property Phones As New List(Of String)()
        Public Property RawJson As JObject
    End Class

    ''' <summary>Thrown for Graph errors that callers should surface to the user (auth, 403, throttling).</summary>
    Public Class M365GraphException
        Inherits Exception
        Public Property StatusCode As Integer
        Public Property Source_ As M365SearchSources
        Public Sub New(message As String, statusCode As Integer, Optional inner As Exception = Nothing)
            MyBase.New(message, inner)
            Me.StatusCode = statusCode
        End Sub
    End Class

End Namespace