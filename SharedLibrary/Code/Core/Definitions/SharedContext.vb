' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: SharedContext.vb
' Purpose: Defines `SharedContext`, a shared, mutable context container that implements
'          `ISharedContext` and is used to pass configuration values and prompt strings
'          through the SharedLibrary and host add-ins.
'
' Architecture / How it works:
'  - `ISharedContext` defines the complete contract (properties) for shared settings and
'    runtime values that are initialized by each host project (e.g., in `ThisAddIn.Properties.vb`)
'    and then read/updated by SharedLibrary code.
'  - `SharedContext` provides the concrete implementation via auto-properties, making the
'    context easy to construct, initialize, and pass around as a single object.
'  - `SharedContext.New()` initializes the two list properties used as an in-memory prompt
'    library (`PromptTitles` and `PromptLibrary`).
'
' Maintenance:
'  - When adding a new shared property, update both `ISharedContext` and the corresponding
'    `SharedContext` implementation, and ensure each host project initializes it consistently
'    (see the existing checklist below).
' =============================================================================

Option Strict On
Option Explicit On

Namespace SharedLibrary
    Partial Public Class SharedContext

        Implements ISharedContext

        ' To add new Shared Properties (e.g., common configuration variable such as INI_Model or system prompts such as SP_Explain):
        ' 
        ' 1. Add Property declaration to ISharedContext interface below
        ' 2. Implement Property below  (with Implements ISharedContext.PropertyName)
        ' 3. Update for each project the ThisAddIn.Properties.vb file to initialize the Property
        ' 4. Search and update SharedLibrary / Core / Config / ShareMethods.LoadConfig.vb (search for a pre-existing Property usage)
        ' 5. Search and update SharedLibrary / Core / Config / ShareMethods.Settings.vb (search for a pre-existing Property usage)
        ' 6. Search and update SharedLibrary / Core / Definitions / SharedMethods.Constants.vb (if you need to update Property usage)
        ' 7. Add a default value in Sharedlibrary / Core / Constants if there is one and include it in the exclusion list of UpdateAppConfig() (no one needed for Boolean = False or empty strings)
        ' 8. Update the IsBooleanSetting() in SharedLibrary / Core / Config / ShareMethods.Settings.vb

        Public Interface ISharedContext

            Property INI_APIKey As String
            Property INI_APIKeyBack As String
            Property INI_Temperature As String
            Property INI_Timeout As Long
            Property INI_MaxOutputToken As Integer
            Property INI_Model As String
            Property INI_Endpoint As String
            Property INI_HeaderA As String
            Property INI_HeaderB As String
            Property INI_APICall As String
            Property INI_APICall_Object As String
            Property INI_Response As String
            Property INI_Anon As String
            Property INI_TokenCount As String
            Property INI_DoubleS As Boolean
            Property INI_Clean As Boolean
            Property INI_Ignore As Boolean
            Property INI_Location As String
            Property INI_NoDash As Boolean
            Property INI_MarkdownBubbles As Boolean
            Property INI_PreCorrection As String
            Property INI_PostCorrection As String
            Property INI_APIEncrypted As Boolean
            Property INI_APIKeyPrefix As String
            Property INI_MarkupMethodOutlook As Integer
            Property INI_MarkupDiffCap As Integer
            Property INI_MarkupRegexCap As Integer
            Property INI_OpenSSLPath As String
            Property INI_OAuth2 As Boolean
            Property INI_OAuth2ClientMail As String
            Property INI_OAuth2Scopes As String
            Property INI_OAuth2Endpoint As String
            Property INI_OAuth2ATExpiry As Long
            Property INI_SecondAPI As Boolean
            Property INI_APIKey_2 As String
            Property INI_APIKeyBack_2 As String
            Property INI_Temperature_2 As String
            Property INI_Timeout_2 As Long
            Property INI_MaxOutputToken_2 As Integer
            Property INI_Model_2 As String
            Property INI_Endpoint_2 As String
            Property INI_HeaderA_2 As String
            Property INI_HeaderB_2 As String
            Property INI_APICall_2 As String
            Property INI_APICall_Object_2 As String
            Property INI_Response_2 As String
            Property INI_Anon_2 As String
            Property INI_TokenCount_2 As String
            Property INI_APIEncrypted_2 As Boolean
            Property INI_APIKeyPrefix_2 As String
            Property INI_OAuth2_2 As Boolean
            Property INI_OAuth2ClientMail_2 As String
            Property INI_OAuth2Scopes_2 As String
            Property INI_OAuth2Endpoint_2 As String
            Property INI_OAuth2ATExpiry_2 As Long
            Property INI_APIDebug As Boolean
            Property INI_UsageRestrictions As String
            Property INI_LogPath As String
            Property INI_Language1 As String
            Property INI_Language2 As String
            Property INI_DefaultPrefix As String
            Property INI_KeepFormat1 As Boolean
            Property INI_KeepFormat2 As Boolean
            Property INI_KeepFormatCap As Integer
            Property INI_MarkdownConvert As Boolean
            Property INI_KeepParaFormatInline As Boolean
            Property INI_ReplaceText1 As Boolean
            Property INI_ReplaceText2 As Boolean
            Property INI_ReplaceText2Override As String
            Property INI_DoMarkupOutlook As Boolean
            Property INI_DoMarkupWord As Boolean
            Property INI_RoastMe As Boolean
            Property DecodedAPI As String
            Property DecodedAPI_2 As String
            Property TokenExpiry As DateTime
            Property TokenExpiry_2 As DateTime
            Property Codebasis As String
            Property GPTSetupError As Boolean
            Property INIloaded As Boolean
            Property RDV As String
            Property InitialConfigFailed As Boolean
            Property INI_ContextMenu As Boolean
            Property INI_NoLocalConfig As Boolean
            Property INI_UpdateCheckInterval As Integer
            Property INI_UpdatePath As String
            Property INI_HelpMeInkyPath As String
            Property INI_DiscussInkyPath As String
            Property INI_DiscussInkyPathLocal As String
            Property INI_ExtractorPath As String
            Property INI_ExtractorPathLocal As String
            Property INI_RenameLibPath As String
            Property INI_RenameLibPathLocal As String
            Property INI_RedactionInstructionsPath As String
            Property INI_RedactionInstructionsPathLocal As String
            Property INI_SpeechModelPath As String
            Property INI_LocalModelPath As String
            Property INI_TTSEndpoint As String
            Property SP_Translate As String
            Property SP_Translate_Document As String
            Property SP_Correct As String

            Property SP_Correct_Document As String
            Property SP_Improve As String
            Property SP_Explain As String
            Property SP_FindClause As String
            Property SP_FindClause_Clean As String
            Property SP_ApplyDocStyle As String
            Property SP_ApplyDocStyle_NumberingHint As String
            Property SP_DocCheck_Clause As String
            Property SP_DocCheck_MultiClause As String
            Property SP_DocCheck_MultiClauseSum As String
            Property SP_DocCheck_MultiClauseSum_Bubbles As String
            Property SP_SuggestTitles As String
            Property SP_Friendly As String
            Property SP_Convincing As String
            Property SP_NoFillers As String
            Property SP_Podcast As String
            Property SP_MyStyle_Word As String
            Property SP_MyStyle_Outlook As String
            Property SP_MyStyle_Apply As String
            Property SP_Shorten As String
            Property SP_Filibuster As String
            Property SP_ArgueAgainst As String
            Property SP_InsertClipboard As String
            Property SP_Summarize As String
            Property SP_Markup As String
            Property SP_MailReply As String
            Property SP_MailSumup As String
            Property SP_MailSumup2 As String
            Property SP_FreestyleText As String
            Property SP_FreestyleNoText As String
            Property SP_Freestyle_Document As String
            Property SP_SwitchParty As String
            Property SP_Anonymize As String
            Property SP_Extract As String
            Property SP_ExtractSchema As String
            Property SP_MergeDateRows As String
            Property SP_Rename As String
            Property SP_Redact As String
            Property SP_CheckforII As String
            Property SP_ContextSearch As String
            Property SP_ContextSearchMulti As String
            Property SP_WriteNeatly As String
            Property SP_RangeOfCells As String
            Property SP_ParseFile As String
            Property SP_Ignore As String
            Property SP_Add_Tooling As String
            Property SP_Add_KeepFormulasIntact As String
            Property SP_Add_KeepHTMLIntact As String
            Property SP_Add_KeepInlineIntact As String
            Property SP_Add_Bubbles As String
            Property SP_Add_BubblesExtract As String
            Property SP_Add_BubblesReply As String
            Property SP_Add_Bubbles_Format As String
            Property SP_Add_Batch As String
            Property SP_Add_Slides As String
            Property SP_BubblesExcel As String
            Property SP_Add_Revisions As String
            Property SP_MarkupRegex As String
            Property SP_ChatWord As String
            Property SP_Chat As String
            Property SP_HelpMe As String
            Property SP_DiscussThis_SortOut As String
            Property SP_DiscussThis_SumUp As String
            Property SP_Add_ChatWord_Commands As String
            Property SP_Add_Chat_NoCommands As String
            Property SP_ChatExcel As String
            Property SP_Add_ChatExcel_Commands As String
            Property INI_ChatCap As Integer
            Property INI_ISearch As Boolean
            Property INI_ISearch_Approve As Boolean
            Property INI_ISearch_URL As String
            Property INI_ISearch_ResponseURLStart As String
            Property INI_ISearch_ResponseMask1 As String
            Property INI_ISearch_ResponseMask2 As String
            Property INI_ISearch_Name As String
            Property INI_ISearch_Tries As Integer
            Property INI_ISearch_Results As Integer
            Property INI_ISearch_MaxDepth As Integer
            Property INI_ISearch_Timeout As Long
            Property INI_ISearch_SearchTerm_SP As String
            Property INI_ISearch_Apply_SP_Markup As String
            Property INI_ISearch_Apply_SP As String
            Property INI_Lib As Boolean
            Property INI_Lib_File As String
            Property INI_Lib_Timeout As Long
            Property INI_Lib_Find_SP As String
            Property INI_Lib_Apply_SP As String
            Property INI_Lib_Apply_SP_Markup As String
            Property INI_MarkupMethodHelper As Integer
            Property INI_MarkupMethodWord As Integer
            Property INI_MarkupMethodWordOverride As String
            Property INI_MarkupMethodOutlookOverride As String
            Property INI_ShortcutsWordExcel As String
            Property INI_PromptLib As Boolean
            Property INI_PromptLibPath As String
            Property INI_PromptLibPathLocal As String
            Property INI_MyStylePath As String
            Property INI_AlternateModelPath As String
            Property INI_SpecialServicePath As String
            Property INI_FindClausePath As String
            Property INI_FindClausePathLocal As String
            Property INI_WebAgentPath As String
            Property INI_WebAgentPathLocal As String
            Property INI_DocCheckPath As String
            Property INI_DocCheckPathLocal As String
            Property INI_DocStylePath As String
            Property INI_DocStylePathLocal As String
            Property INI_PromptLibPath_Transcript As String
            Property PromptLibrary() As List(Of String)
            Property PromptTitles() As List(Of String)
            Property MenusAdded As Boolean
            Property INI_Model_Parameter1 As String
            Property INI_Model_Parameter2 As String
            Property INI_Model_Parameter3 As String
            Property INI_Model_Parameter4 As String
            Property SP_FindPrompts As String
            Property SP_MergePrompt As String
            Property SP_MergePrompt2 As String
            Property SP_Add_MergePrompt As String
            Property Ignore As String
            Property Location As String

            Property INI_NoHelperDownload As Boolean

            ' Master switch for INI update mechanism
            Property INI_UpdateIni As Boolean
            ' Clients that are permitted to run updates from remote sources
            Property INI_UpdateIniAllowRemote As Boolean
            ' Skip signature verification if True
            Property INI_UpdateIniNoSignature As Boolean
            ' Update source for redink.ini: "path; keylist; base64_public_key"
            Property INI_UpdateSource As String
            ' Override ignore settings with file-specific and segment-specific rules
            Property INI_UpdateIniClients As String
            ' Allow HTTPS sources (vs local/network only)
            Property INI_UpdateIniIgnoreOverride As String
            ' Silent update mode: Controls whether updates are applied without user interaction
            Property INI_UpdateIniSilentMode As Integer
            ' Log silent update actions to a file for audit purposes
            Property INI_UpdateIniSilentLog As Boolean


            ' Tooling / tool-call settings 
            Property INI_ToolingLogWindow As Boolean
            Property INI_ToolingDryRun As Boolean
            Property INI_ToolingMaximumIterations As Integer

            Property INI_APICall_ToolInstructions_2 As String
            Property INI_APICall_ToolInstructions_Template_2 As String
            Property INI_APICall_ToolResponses_2 As String
            Property INI_APICall_ToolResponses_Template_2 As String
            Property INI_APICall_ToolCallPart_Template_2 As String
            Property INI_ToolCallDetectionPattern_2 As String
            Property INI_ToolCallExtractionMap_2 As String

            Property INI_LogoPathLarge As String
            Property INI_LogoPathMedium As String
            Property INI_LogoPath As String
            Property INI_BrandingName As String


        End Interface

        Public Sub New()
            ' Initialize the PromptTitles and PromptLibrary properties
            PromptTitles = New List(Of String)()
            PromptLibrary = New List(Of String)()
        End Sub

        Public Property INI_APIKey As String Implements ISharedContext.INI_APIKey
        Public Property INI_APIKeyBack As String Implements ISharedContext.INI_APIKeyBack
        Public Property INI_Temperature As String Implements ISharedContext.INI_Temperature
        Public Property INI_Timeout As Long Implements ISharedContext.INI_Timeout
        Public Property INI_MaxOutputToken As Integer Implements ISharedContext.INI_MaxOutputToken
        Public Property INI_Model As String Implements ISharedContext.INI_Model
        Public Property INI_Endpoint As String Implements ISharedContext.INI_Endpoint
        Public Property INI_HeaderA As String Implements ISharedContext.INI_HeaderA
        Public Property INI_HeaderB As String Implements ISharedContext.INI_HeaderB
        Public Property INI_APICall As String Implements ISharedContext.INI_APICall
        Public Property INI_APICall_Object As String Implements ISharedContext.INI_APICall_Object
        Public Property INI_Response As String Implements ISharedContext.INI_Response
        Public Property INI_Anon As String Implements ISharedContext.INI_Anon
        Public Property INI_TokenCount As String Implements ISharedContext.INI_TokenCount
        Public Property INI_DoubleS As Boolean Implements ISharedContext.INI_DoubleS
        Public Property INI_Clean As Boolean Implements ISharedContext.INI_Clean
        Public Property INI_Ignore As Boolean Implements ISharedContext.INI_Ignore
        Public Property INI_Location As String Implements ISharedContext.INI_Location
        Public Property INI_NoDash As Boolean Implements ISharedContext.INI_NoDash
        Public Property INI_MarkdownBubbles As Boolean Implements ISharedContext.INI_MarkdownBubbles
        Public Property INI_PreCorrection As String Implements ISharedContext.INI_PreCorrection
        Public Property INI_PostCorrection As String Implements ISharedContext.INI_PostCorrection
        Public Property INI_APIEncrypted As Boolean Implements ISharedContext.INI_APIEncrypted
        Public Property INI_APIKeyPrefix As String Implements ISharedContext.INI_APIKeyPrefix
        Public Property INI_MarkupMethodOutlook As Integer Implements ISharedContext.INI_MarkupMethodOutlook
        Public Property INI_MarkupDiffCap As Integer Implements ISharedContext.INI_MarkupDiffCap
        Public Property INI_MarkupRegexCap As Integer Implements ISharedContext.INI_MarkupRegexCap
        Public Property INI_OpenSSLPath As String Implements ISharedContext.INI_OpenSSLPath
        Public Property INI_OAuth2 As Boolean Implements ISharedContext.INI_OAuth2
        Public Property INI_OAuth2ClientMail As String Implements ISharedContext.INI_OAuth2ClientMail
        Public Property INI_OAuth2Scopes As String Implements ISharedContext.INI_OAuth2Scopes
        Public Property INI_OAuth2Endpoint As String Implements ISharedContext.INI_OAuth2Endpoint
        Public Property INI_OAuth2ATExpiry As Long Implements ISharedContext.INI_OAuth2ATExpiry
        Public Property INI_SecondAPI As Boolean Implements ISharedContext.INI_SecondAPI
        Public Property INI_APIKey_2 As String Implements ISharedContext.INI_APIKey_2
        Public Property INI_APIKeyBack_2 As String Implements ISharedContext.INI_APIKeyBack_2
        Public Property INI_Temperature_2 As String Implements ISharedContext.INI_Temperature_2
        Public Property INI_Timeout_2 As Long Implements ISharedContext.INI_Timeout_2
        Public Property INI_MaxOutputToken_2 As Integer Implements ISharedContext.INI_MaxOutputToken_2
        Public Property INI_Model_2 As String Implements ISharedContext.INI_Model_2
        Public Property INI_Endpoint_2 As String Implements ISharedContext.INI_Endpoint_2
        Public Property INI_HeaderA_2 As String Implements ISharedContext.INI_HeaderA_2
        Public Property INI_HeaderB_2 As String Implements ISharedContext.INI_HeaderB_2
        Public Property INI_APICall_2 As String Implements ISharedContext.INI_APICall_2
        Public Property INI_APICall_Object_2 As String Implements ISharedContext.INI_APICall_Object_2
        Public Property INI_Response_2 As String Implements ISharedContext.INI_Response_2
        Public Property INI_Anon_2 As String Implements ISharedContext.INI_Anon_2
        Public Property INI_TokenCount_2 As String Implements ISharedContext.INI_TokenCount_2
        Public Property INI_APIEncrypted_2 As Boolean Implements ISharedContext.INI_APIEncrypted_2
        Public Property INI_APIKeyPrefix_2 As String Implements ISharedContext.INI_APIKeyPrefix_2
        Public Property INI_OAuth2_2 As Boolean Implements ISharedContext.INI_OAuth2_2
        Public Property INI_OAuth2ClientMail_2 As String Implements ISharedContext.INI_OAuth2ClientMail_2
        Public Property INI_OAuth2Scopes_2 As String Implements ISharedContext.INI_OAuth2Scopes_2
        Public Property INI_OAuth2Endpoint_2 As String Implements ISharedContext.INI_OAuth2Endpoint_2
        Public Property INI_OAuth2ATExpiry_2 As Long Implements ISharedContext.INI_OAuth2ATExpiry_2
        Public Property INI_APIDebug As Boolean Implements ISharedContext.INI_APIDebug
        Public Property INI_UsageRestrictions As String Implements ISharedContext.INI_UsageRestrictions
        Public Property INI_LogPath As String Implements ISharedContext.INI_LogPath
        Public Property INI_Language1 As String Implements ISharedContext.INI_Language1
        Public Property INI_Language2 As String Implements ISharedContext.INI_Language2
        Public Property INI_MarkdownConvert As Boolean Implements ISharedContext.INI_MarkdownConvert
        Public Property INI_DefaultPrefix As String Implements ISharedContext.INI_DefaultPrefix
        Public Property INI_KeepFormat1 As Boolean Implements ISharedContext.INI_KeepFormat1
        Public Property INI_KeepFormat2 As Boolean Implements ISharedContext.INI_KeepFormat2
        Public Property INI_KeepFormatCap As Integer Implements ISharedContext.INI_KeepFormatCap
        Public Property INI_KeepParaFormatInline As Boolean Implements ISharedContext.INI_KeepParaFormatInline
        Public Property INI_ReplaceText1 As Boolean Implements ISharedContext.INI_ReplaceText1
        Public Property INI_ReplaceText2 As Boolean Implements ISharedContext.INI_ReplaceText2
        Public Property INI_ReplaceText2Override As String Implements ISharedContext.INI_ReplaceText2Override
        Public Property INI_DoMarkupOutlook As Boolean Implements ISharedContext.INI_DoMarkupOutlook
        Public Property INI_DoMarkupWord As Boolean Implements ISharedContext.INI_DoMarkupWord
        Public Property INI_RoastMe As Boolean Implements ISharedContext.INI_RoastMe
        Public Property DecodedAPI As String Implements ISharedContext.DecodedAPI
        Public Property DecodedAPI_2 As String Implements ISharedContext.DecodedAPI_2
        Public Property TokenExpiry As DateTime Implements ISharedContext.TokenExpiry
        Public Property TokenExpiry_2 As DateTime Implements ISharedContext.TokenExpiry_2
        Public Property Codebasis As String Implements ISharedContext.Codebasis

        Public Property GPTSetupError As Boolean Implements ISharedContext.GPTSetupError
        Public Property INIloaded As Boolean Implements ISharedContext.INIloaded
        Public Property RDV As String Implements ISharedContext.RDV
        Public Property InitialConfigFailed As Boolean Implements ISharedContext.InitialConfigFailed
        Public Property INI_ContextMenu As Boolean Implements ISharedContext.INI_ContextMenu
        Public Property INI_NoLocalConfig As Boolean Implements ISharedContext.INI_NoLocalConfig
        Public Property INI_UpdateCheckInterval As Integer Implements ISharedContext.INI_UpdateCheckInterval
        Public Property INI_UpdatePath As String Implements ISharedContext.INI_UpdatePath
        Public Property INI_HelpMeInkyPath As String Implements ISharedContext.INI_HelpMeInkyPath
        Public Property INI_DiscussInkyPath As String Implements ISharedContext.INI_DiscussInkyPath
        Public Property INI_DiscussInkyPathLocal As String Implements ISharedContext.INI_DiscussInkyPathLocal
        Public Property INI_ExtractorPath As String Implements ISharedContext.INI_ExtractorPath
        Public Property INI_ExtractorPathLocal As String Implements ISharedContext.INI_ExtractorPathLocal
        Public Property INI_RenameLibPath As String Implements ISharedContext.INI_RenameLibPath
        Public Property INI_RenameLibPathLocal As String Implements ISharedContext.INI_RenameLibPathLocal

        Public Property INI_RedactionInstructionsPath As String Implements ISharedContext.INI_RedactionInstructionsPath
        Public Property INI_RedactionInstructionsPathLocal As String Implements ISharedContext.INI_RedactionInstructionsPathLocal
        Public Property INI_SpeechModelPath As String Implements ISharedContext.INI_SpeechModelPath
        Public Property INI_LocalModelPath As String Implements ISharedContext.INI_LocalModelPath
        Public Property INI_TTSEndpoint As String Implements ISharedContext.INI_TTSEndpoint
        Public Property SP_Translate As String Implements ISharedContext.SP_Translate
        Public Property SP_Translate_Document As String Implements ISharedContext.SP_Translate_Document
        Public Property SP_Correct As String Implements ISharedContext.SP_Correct
        Public Property SP_Correct_Document As String Implements ISharedContext.SP_Correct_Document
        Public Property SP_Improve As String Implements ISharedContext.SP_Improve
        Public Property SP_Explain As String Implements ISharedContext.SP_Explain
        Public Property SP_FindClause As String Implements ISharedContext.SP_FindClause
        Public Property SP_FindClause_Clean As String Implements ISharedContext.SP_FindClause_Clean
        Public Property SP_ApplyDocStyle As String Implements ISharedContext.SP_ApplyDocStyle
        Public Property SP_ApplyDocStyle_NumberingHint As String Implements ISharedContext.SP_ApplyDocStyle_NumberingHint
        Public Property SP_DocCheck_Clause As String Implements ISharedContext.SP_DocCheck_Clause
        Public Property SP_DocCheck_MultiClause As String Implements ISharedContext.SP_DocCheck_MultiClause
        Public Property SP_DocCheck_MultiClauseSum As String Implements ISharedContext.SP_DocCheck_MultiClauseSum
        Public Property SP_DocCheck_MultiClauseSum_Bubbles As String Implements ISharedContext.SP_DocCheck_MultiClauseSum_Bubbles
        Public Property SP_SuggestTitles As String Implements ISharedContext.SP_SuggestTitles
        Public Property SP_Friendly As String Implements ISharedContext.SP_Friendly
        Public Property SP_Convincing As String Implements ISharedContext.SP_Convincing
        Public Property SP_NoFillers As String Implements ISharedContext.SP_NoFillers
        Public Property SP_Podcast As String Implements ISharedContext.SP_Podcast
        Public Property SP_MyStyle_Word As String Implements ISharedContext.SP_MyStyle_Word
        Public Property SP_MyStyle_Outlook As String Implements ISharedContext.SP_MyStyle_Outlook
        Public Property SP_MyStyle_Apply As String Implements ISharedContext.SP_MyStyle_Apply

        Public Property SP_Shorten As String Implements ISharedContext.SP_Shorten

        Public Property SP_Filibuster As String Implements ISharedContext.SP_Filibuster
        Public Property SP_ArgueAgainst As String Implements ISharedContext.SP_ArgueAgainst
        Public Property SP_InsertClipboard As String Implements ISharedContext.SP_InsertClipboard
        Public Property SP_Summarize As String Implements ISharedContext.SP_Summarize
        Public Property SP_Markup As String Implements ISharedContext.SP_Markup

        Public Property SP_MailReply As String Implements ISharedContext.SP_MailReply
        Public Property SP_MailSumup As String Implements ISharedContext.SP_MailSumup
        Public Property SP_MailSumup2 As String Implements ISharedContext.SP_MailSumup2
        Public Property SP_FreestyleText As String Implements ISharedContext.SP_FreestyleText
        Public Property SP_FreestyleNoText As String Implements ISharedContext.SP_FreestyleNoText
        Public Property SP_Freestyle_Document As String Implements ISharedContext.SP_Freestyle_Document
        Public Property SP_SwitchParty As String Implements ISharedContext.SP_SwitchParty
        Public Property SP_Anonymize As String Implements ISharedContext.SP_Anonymize
        Public Property SP_Extract As String Implements ISharedContext.SP_Extract
        Public Property SP_ExtractSchema As String Implements ISharedContext.SP_ExtractSchema

        Public Property SP_MergeDateRows As String Implements ISharedContext.SP_MergeDateRows
        Public Property SP_Rename As String Implements ISharedContext.SP_Rename
        Public Property SP_Redact As String Implements ISharedContext.SP_Redact
        Public Property SP_CheckforII As String Implements ISharedContext.SP_CheckforII
        Public Property SP_ContextSearch As String Implements ISharedContext.SP_ContextSearch
        Public Property SP_ContextSearchMulti As String Implements ISharedContext.SP_ContextSearchMulti
        Public Property SP_RangeOfCells As String Implements ISharedContext.SP_RangeOfCells
        Public Property SP_ParseFile As String Implements ISharedContext.SP_ParseFile
        Public Property SP_Ignore As String Implements ISharedContext.SP_Ignore
        Public Property SP_WriteNeatly As String Implements ISharedContext.SP_WriteNeatly
        Public Property SP_Add_Tooling As String Implements ISharedContext.SP_Add_Tooling
        Public Property SP_Add_KeepFormulasIntact As String Implements ISharedContext.SP_Add_KeepFormulasIntact
        Public Property SP_Add_KeepHTMLIntact As String Implements ISharedContext.SP_Add_KeepHTMLIntact
        Public Property SP_Add_KeepInlineIntact As String Implements ISharedContext.SP_Add_KeepInlineIntact
        Public Property SP_Add_Bubbles As String Implements ISharedContext.SP_Add_Bubbles
        Public Property SP_Add_BubblesExtract As String Implements ISharedContext.SP_Add_BubblesExtract
        Public Property SP_Add_BubblesReply As String Implements ISharedContext.SP_Add_BubblesReply
        Public Property SP_Add_Bubbles_Format As String Implements ISharedContext.SP_Add_Bubbles_Format
        Public Property SP_Add_Batch As String Implements ISharedContext.SP_Add_Batch
        Public Property SP_Add_Slides As String Implements ISharedContext.SP_Add_Slides
        Public Property SP_BubblesExcel As String Implements ISharedContext.SP_BubblesExcel
        Public Property SP_Add_Revisions As String Implements ISharedContext.SP_Add_Revisions
        Public Property SP_MarkupRegex As String Implements ISharedContext.SP_MarkupRegex
        Public Property SP_ChatWord As String Implements ISharedContext.SP_ChatWord

        Public Property SP_Chat As String Implements ISharedContext.SP_Chat
        Public Property SP_HelpMe As String Implements ISharedContext.SP_HelpMe
        Public Property SP_DiscussThis_SortOut As String Implements ISharedContext.SP_DiscussThis_SortOut
        Public Property SP_DiscussThis_SumUp As String Implements ISharedContext.SP_DiscussThis_SumUp
        Public Property SP_Add_ChatWord_Commands As String Implements ISharedContext.SP_Add_ChatWord_Commands
        Public Property SP_Add_Chat_NoCommands As String Implements ISharedContext.SP_Add_Chat_NoCommands
        Public Property SP_ChatExcel As String Implements ISharedContext.SP_ChatExcel
        Public Property SP_Add_ChatExcel_Commands As String Implements ISharedContext.SP_Add_ChatExcel_Commands
        Public Property INI_ChatCap As Integer Implements ISharedContext.INI_ChatCap
        Public Property INI_ISearch As Boolean Implements ISharedContext.INI_ISearch
        Public Property INI_ISearch_Approve As Boolean Implements ISharedContext.INI_ISearch_Approve
        Public Property INI_ISearch_URL As String Implements ISharedContext.INI_ISearch_URL
        Public Property INI_ISearch_ResponseURLStart As String Implements ISharedContext.INI_ISearch_ResponseURLStart
        Public Property INI_ISearch_ResponseMask1 As String Implements ISharedContext.INI_ISearch_ResponseMask1
        Public Property INI_ISearch_ResponseMask2 As String Implements ISharedContext.INI_ISearch_ResponseMask2
        Public Property INI_ISearch_Name As String Implements ISharedContext.INI_ISearch_Name
        Public Property INI_ISearch_Tries As Integer Implements ISharedContext.INI_ISearch_Tries
        Public Property INI_ISearch_Results As Integer Implements ISharedContext.INI_ISearch_Results
        Public Property INI_ISearch_MaxDepth As Integer Implements ISharedContext.INI_ISearch_MaxDepth
        Public Property INI_ISearch_Timeout As Long Implements ISharedContext.INI_ISearch_Timeout
        Public Property INI_ISearch_SearchTerm_SP As String Implements ISharedContext.INI_ISearch_SearchTerm_SP
        Public Property INI_ISearch_Apply_SP As String Implements ISharedContext.INI_ISearch_Apply_SP
        Public Property INI_ISearch_Apply_SP_Markup As String Implements ISharedContext.INI_ISearch_Apply_SP_Markup
        Public Property INI_Lib As Boolean Implements ISharedContext.INI_Lib
        Public Property INI_Lib_File As String Implements ISharedContext.INI_Lib_File
        Public Property INI_Lib_Timeout As Long Implements ISharedContext.INI_Lib_Timeout
        Public Property INI_Lib_Find_SP As String Implements ISharedContext.INI_Lib_Find_SP
        Public Property INI_Lib_Apply_SP_Markup As String Implements ISharedContext.INI_Lib_Apply_SP_Markup
        Public Property INI_Lib_Apply_SP As String Implements ISharedContext.INI_Lib_Apply_SP
        Public Property INI_MarkupMethodHelper As Integer Implements ISharedContext.INI_MarkupMethodHelper
        Public Property INI_MarkupMethodWord As Integer Implements ISharedContext.INI_MarkupMethodWord
        Public Property INI_MarkupMethodWordOverride As String Implements ISharedContext.INI_MarkupMethodWordOverride

        Public Property INI_MarkupMethodOutlookOverride As String Implements ISharedContext.INI_MarkupMethodOutlookOverride
        Public Property INI_ShortcutsWordExcel As String Implements ISharedContext.INI_ShortcutsWordExcel
        Public Property INI_PromptLib As Boolean Implements ISharedContext.INI_PromptLib
        Public Property INI_PromptLibPath As String Implements ISharedContext.INI_PromptLibPath
        Public Property INI_PromptLibPathLocal As String Implements ISharedContext.INI_PromptLibPathLocal
        Public Property INI_MyStylePath As String Implements ISharedContext.INI_MyStylePath
        Public Property INI_AlternateModelPath As String Implements ISharedContext.INI_AlternateModelPath
        Public Property INI_SpecialServicePath As String Implements ISharedContext.INI_SpecialServicePath
        Public Property INI_FindClausePath As String Implements ISharedContext.INI_FindClausePath
        Public Property INI_FindClausePathLocal As String Implements ISharedContext.INI_FindClausePathLocal
        Public Property INI_WebAgentPath As String Implements ISharedContext.INI_WebAgentPath
        Public Property INI_WebAgentPathLocal As String Implements ISharedContext.INI_WebAgentPathLocal
        Public Property INI_DocCheckPath As String Implements ISharedContext.INI_DocCheckPath
        Public Property INI_DocCheckPathLocal As String Implements ISharedContext.INI_DocCheckPathLocal
        Public Property INI_DocStylePath As String Implements ISharedContext.INI_DocStylePath
        Public Property INI_DocStylePathLocal As String Implements ISharedContext.INI_DocStylePathLocal
        Public Property INI_PromptLibPath_Transcript As String Implements ISharedContext.INI_PromptLibPath_Transcript
        Public Property PromptLibrary() As List(Of String) Implements ISharedContext.PromptLibrary
        Public Property PromptTitles() As List(Of String) Implements ISharedContext.PromptTitles
        Public Property MenusAdded As Boolean Implements ISharedContext.MenusAdded
        Public Property INI_Model_Parameter1 As String Implements ISharedContext.INI_Model_Parameter1
        Public Property INI_Model_Parameter2 As String Implements ISharedContext.INI_Model_Parameter2
        Public Property INI_Model_Parameter3 As String Implements ISharedContext.INI_Model_Parameter3
        Public Property INI_Model_Parameter4 As String Implements ISharedContext.INI_Model_Parameter4

        Public Property SP_FindPrompts As String Implements ISharedContext.SP_FindPrompts
        Public Property SP_MergePrompt As String Implements ISharedContext.SP_MergePrompt
        Public Property SP_MergePrompt2 As String Implements ISharedContext.SP_MergePrompt2
        Public Property SP_Add_MergePrompt As String Implements ISharedContext.SP_Add_MergePrompt

        Public Property INI_NoHelperDownload As Boolean Implements ISharedContext.INI_NoHelperDownload
        Public Property INI_UpdateIni As Boolean Implements ISharedContext.INI_UpdateIni
        Public Property INI_UpdateIniClients As String Implements ISharedContext.INI_UpdateIniClients
        Public Property INI_UpdateIniAllowRemote As Boolean Implements ISharedContext.INI_UpdateIniAllowRemote
        Public Property INI_UpdateIniNoSignature As Boolean Implements ISharedContext.INI_UpdateIniNoSignature
        Public Property INI_UpdateSource As String Implements ISharedContext.INI_UpdateSource
        Public Property INI_UpdateIniIgnoreOverride As String Implements ISharedContext.INI_UpdateIniIgnoreOverride
        Public Property INI_UpdateIniSilentMode As Integer Implements ISharedContext.INI_UpdateIniSilentMode
        Public Property INI_UpdateIniSilentLog As Boolean Implements ISharedContext.INI_UpdateIniSilentLog
        Public Property Ignore As String Implements ISharedContext.Ignore
        Public Property Location As String Implements ISharedContext.Location


        ' Tooling / tool-call settings 
        Public Property INI_ToolingLogWindow As Boolean Implements ISharedContext.INI_ToolingLogWindow
        Public Property INI_ToolingDryRun As Boolean Implements ISharedContext.INI_ToolingDryRun
        Public Property INI_ToolingMaximumIterations As Integer Implements ISharedContext.INI_ToolingMaximumIterations
        Public Property INI_APICall_ToolInstructions_2 As String Implements ISharedContext.INI_APICall_ToolInstructions_2
        Public Property INI_APICall_ToolInstructions_Template_2 As String Implements ISharedContext.INI_APICall_ToolInstructions_Template_2
        Public Property INI_APICall_ToolResponses_2 As String Implements ISharedContext.INI_APICall_ToolResponses_2
        Public Property INI_APICall_ToolResponses_Template_2 As String Implements ISharedContext.INI_APICall_ToolResponses_Template_2
        Public Property INI_APICall_ToolCallPart_Template_2 As String Implements ISharedContext.INI_APICall_ToolCallPart_Template_2
        Public Property INI_ToolCallDetectionPattern_2 As String Implements ISharedContext.INI_ToolCallDetectionPattern_2
        Public Property INI_ToolCallExtractionMap_2 As String Implements ISharedContext.INI_ToolCallExtractionMap_2

        Public Property INI_LogoPathLarge As String Implements ISharedContext.INI_LogoPathLarge
        Public Property INI_LogoPathMedium As String Implements ISharedContext.INI_LogoPathMedium
        Public Property INI_LogoPath As String Implements ISharedContext.INI_LogoPath
        Public Property INI_BrandingName As String Implements ISharedContext.INI_BrandingName


    End Class
End Namespace