' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.AutoPilot.Scan.Contracts.vb
' Purpose: Scan inbox contracts; no model-controlled transport.
'
' Architecture:
'  - Part of the AutoPilot Scan Handler; configuration and recipient authority
'    remain in the validated, persisted scan register.
'  - Document/model content never grants transport or filesystem permissions.
'  - Host lifecycle, shared UI and retained operation checkpoints are reused.
' =============================================================================
Option Explicit On
Option Strict On
Option Infer On

Namespace ScanInbox
    Public Class Configuration
        Public Property SchemaVersion As System.Int32 = 1
        Public Property Revision As System.Int64
        Public Property InitialLookbackDays As System.Int32 = 7
        Public Property OverlapHours As System.Int32 = 24
        Public Property Recipients As New System.Collections.Generic.List(Of RegisteredRecipient)()
        Public Property RuleSets As New System.Collections.Generic.List(Of RuleSet)()
    End Class

    Public Class RegisteredRecipient
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Active As System.Boolean = True
        Public Property Name As System.String = ""
        Public Property Aliases As New System.Collections.Generic.List(Of System.String)()
        Public Property Organization As System.String = ""
        Public Property Location As System.String = ""
        Public Property Attributes As New System.Collections.Generic.List(Of System.String)()
        Public Property SmtpAddress As System.String = ""
        Public Property Instruction As System.String = ""
    End Class

    Public Class RuleSet
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Name As System.String = ""
        Public Property Active As System.Boolean = True
        Public Property ScannerSenders As New System.Collections.Generic.List(Of System.String)()
        Public Property Mailbox As System.String = ""
        Public Property TimeZoneId As System.String = System.TimeZoneInfo.Local.Id
        ' Explicitly supported profile: calendar days; other legal rules require review.
        Public Property DeadlineProfile As System.String = "calendar_days"
        Public Property ReminderDays As New System.Collections.Generic.List(Of System.Int32) From {10, 30}
        Public Property RetentionDays As System.Int32 = 30
        Public Property RenameAttachments As System.Boolean = True
        Public Property AttachmentNameTemplate As System.String = AttachmentNaming.DefaultTemplate
        Public Property AttachmentNamingInstruction As System.String = ""
        Public Property SeparateMessages As System.Boolean
        Public Property DefaultRecipientId As System.String = ""
        Public Property AllowedRecipientIds As New System.Collections.Generic.List(Of System.String)()
        Public Property AdditionalRules As New System.Collections.Generic.List(Of AdditionalRule)()
        Public Property Guidance As System.String = ""
        Public Property SubjectTemplate As System.String = ""
    End Class

    Public Class AdditionalRule
        Public Property Id As System.String = System.Guid.NewGuid().ToString("N")
        Public Property Active As System.Boolean = True
        ' exact: bounded literal match, never a regex supplied by a document/model.
        Public Property Kind As System.String = "exact"
        Public Property Criteria As System.String = ""
        Public Property RecipientIds As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public Class Document
        Public Property Id As System.String = ""
        Public Property FileName As System.String = ""
        Public Property RelativePath As System.String = ""
        Public Property Sha256 As System.String = ""
        Public Property Pages As New System.Collections.Generic.List(Of System.String)()
        Public Property Complete As System.Boolean
        Public Property Problems As New System.Collections.Generic.List(Of System.String)()
    End Class

    ' Every substantive claim carries attachment-relative evidence; ISO dates only.
    Public Class Evidence
        Public Property DocumentId As System.String = ""
        Public Property StartPage As System.Int32
        Public Property EndPage As System.Int32
        Public Property Quote As System.String = ""
    End Class

    Public Class Presentation
        Public Property SubjectSummary As System.String = ""
        Public Property Language As System.String = ""
        Public Property Labels As New System.Collections.Generic.Dictionary(Of System.String, System.String)()
        Public Property EventTitles As New System.Collections.Generic.List(Of System.String)()
        Public Property ReviewNotes As New System.Collections.Generic.List(Of System.String)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Summary As New System.Collections.Generic.List(Of System.String)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Actions As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public Class Extraction
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Sources As New System.Collections.Generic.List(Of Source)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Deadlines As New System.Collections.Generic.List(Of Deadline)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Stamps As New System.Collections.Generic.List(Of ReceiptStamp)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property RuleMatches As New System.Collections.Generic.List(Of RuleMatch)()
        <Newtonsoft.Json.JsonProperty(Required:=Newtonsoft.Json.Required.Always)>
        Public Property Problems As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public Class Source
        Public Property Id As System.String = ""
        Public Property Evidence As New Evidence()
        ' current_cover, current_letter, attachment, historical, unknown
        Public Property FunctionInPackage As System.String = "unknown"
        Public Property DocumentType As System.String = ""
        Public Property Subject As System.String = ""
        Public Property SubjectEvidence As Evidence
        Public Property ContextEvidence As New Evidence()
        Public Property DateIso As System.String = ""
        Public Property DateEvidence As Evidence
        Public Property Certain As System.Boolean
        Public Property Language As System.String = ""
        Public Property Summary As New System.Collections.Generic.List(Of System.String)()
        Public Property Actions As New System.Collections.Generic.List(Of System.String)()
        Public Property Recipients As New System.Collections.Generic.List(Of RecipientClaim)()
        Public Property Relationships As New System.Collections.Generic.List(Of System.String)()
        Public Property Uncertainties As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public Class RecipientClaim
        Public Property CandidateIds As New System.Collections.Generic.List(Of System.String)()
        ' direct, copy, mention, signature; external entities have no candidate ID.
        Public Property Role As System.String = "mention"
        Public Property Name As System.String = ""
        Public Property Organization As System.String = ""
        Public Property Certain As System.Boolean
        Public Property Evidence As New Evidence()
    End Class

    Public Class Deadline
        Public Property Id As System.String = ""
        Public Property Evidence As New Evidence()
        Public Property Action As System.String = ""
        Public Property Person As System.String = ""
        Public Property Status As System.String = "unclear" ' current, historical, unclear
        Public Property CurrentRelation As System.String = ""
        Public Property Certain As System.Boolean
        Public Property DateIso As System.String = ""
        Public Property Duration As System.Int32
        Public Property Unit As System.String = ""
        Public Property Trigger As System.String = "" ' receipt, other, explicit
        Public Property Profile As System.String = ""
    End Class

    Public Class ReceiptStamp
        Public Property DateIso As System.String = ""
        Public Property Relation As System.String = "unclear" ' current, historical, unclear
        Public Property Certain As System.Boolean
        Public Property Evidence As New Evidence()
    End Class

    Public Class RuleMatch
        Public Property RuleId As System.String = ""
        Public Property Certain As System.Boolean
        Public Property Evidence As New Evidence()
    End Class

    Public Class Decision
        Public Property SourceId As System.String = ""
        Public Property UsedDefault As System.Boolean
        Public Property ToIds As New System.Collections.Generic.List(Of System.String)()
        Public Property CcIds As New System.Collections.Generic.List(Of System.String)()
        Public Property DirectIds As New System.Collections.Generic.List(Of System.String)()
        Public Property CopyIds As New System.Collections.Generic.List(Of System.String)()
        Public Property AdditionalIds As New System.Collections.Generic.List(Of System.String)()
        Public Property Reasons As New System.Collections.Generic.List(Of System.String)()
        Public Property ReceiptDate As System.DateTime
        Public Property ReceiptBasis As System.String = ""
        Public Property UsesReceiptStamp As System.Boolean
        Public Property PastDeadlines As New System.Collections.Generic.List(Of CalendarEvent)()
        Public Property Events As New System.Collections.Generic.List(Of CalendarEvent)()
        Public Property Review As New System.Collections.Generic.List(Of System.String)()
    End Class

    Public Class CalendarEvent
        Public Property Id As System.String = ""
        Public Property DateValue As System.DateTime
        Public Property Title As System.String = ""
        Public Property Description As System.String = ""
    End Class

    Public Class Operation
        Public Property SchemaVersion As System.Int32 = 1
        Public Property Id As System.String = ""
        Public Property EntryId As System.String = ""
        Public Property StoreId As System.String = ""
        Public Property ScannerSender As System.String = ""
        Public Property CleanupGroupId As System.String = ""
        Public Property RuleSetId As System.String = ""
        Public Property ConfigurationRevision As System.Int64
        Public Property ConfigurationHash As System.String = ""
        Public Property ReceivedUtc As System.DateTime
        Public Property Mailbox As System.String = ""
        Public Property Subject As System.String = ""
        Public Property Status As System.String = "registered"
        Public Property ErrorMessage As System.String = ""
        Public Property SourceFinalizationStatus As System.String = "pending"
        Public Property ApprovedForProcessing As System.Boolean
        Public Property SpoolCleaned As System.Boolean
        Public Property NextRetryUtc As System.DateTime
        Public Property Documents As New System.Collections.Generic.List(Of Document)()
        Public Property Extraction As Extraction
        Public Property Decision As Decision
        Public Property Presentation As Presentation
        Public Property AttachmentNames As New System.Collections.Generic.List(Of AttachmentName)()
        Public Property Deliveries As New System.Collections.Generic.List(Of Delivery)()
    End Class

    Public Class Delivery
        Public Property Id As System.String = ""
        Public Property Subject As System.String = ""
        Public Property Html As System.String = ""
        Public Property ToIds As New System.Collections.Generic.List(Of System.String)()
        Public Property CcIds As New System.Collections.Generic.List(Of System.String)()
        Public Property RelativeFiles As New System.Collections.Generic.List(Of System.String)()
        ' pending -> submitting (durable before Send) -> submitted; uncertain never blind retry.
        Public Property Status As System.String = "pending"
        Public Property ConversationId As System.String = ""
        Public Property DraftEntryId As System.String = ""
        Public Property DraftStoreId As System.String = ""
        Public Property SubmittedUtc As System.DateTime?
    End Class
End Namespace
