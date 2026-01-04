' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.Properties.vb
' Purpose: Central property bridge exposing configuration, runtime state, prompt
'          templates, and feature flags of the Outlook add‑in through a shared
'          `ISharedContext` instance for consumption by SharedLibrary logic.
'
' Architecture:
' - Backing Context: `_context As ISharedContext` holds all INI / runtime fields; each
'   Shared Property simply forwards get/set to that context (no extra logic here).
' - Primary Model Set: INI_APIKey, INI_APIKeyBack, INI_Temperature, INI_Timeout,
'   INI_MaxOutputToken, INI_Model, INI_Endpoint, INI_HeaderA/B, INI_APICall(+_Object),
'   INI_Response, INI_Anon, INI_TokenCount(+_2), encryption flags and prefixes.
' - Secondary Model Set (alternate / second API): mirrored *_2 properties (keys,
'   temperature, timeout, tokens, model, endpoint, headers, call spec, response,
'   anon, encryption, prefixes, OAuth2 set).
' - Markup / Formatting Controls: INI_DoMarkupOutlook, INI_DoMarkupWord,
'   INI_MarkupMethodOutlook/Word/Helper (+ overrides), INI_MarkupDiffCap,
'   INI_MarkupRegexCap, INI_MarkdownBubbles, INI_KeepFormat1/2, INI_KeepParaFormatInline,
'   INI_KeepFormatCap, INI_ReplaceText1/2 (+ override), SP_MarkupRegex,
'   SP_Add_KeepFormulasIntact / KeepHTMLIntact / KeepInlineIntact.
' - Prompt Templates (system/user additions): Numerous SP_* properties (Translate,
'   Correct, Improve, Explain, FindClause variants, DocCheck variants, SuggestTitles,
'   Style, Shorten, Summarize, Mail reply/sumup, Freestyle, SwitchParty, Anonymize,
'   Redact, CheckforII, ContextSearch variants, RangeOfCells, ParseFile, WriteNeatly,
'   Bubbles (add/extract/reply/format/excel), Batch, Revisions, Chat (Word/Excel generic
'   + additions with or without commands), FindPrompts, MergePrompt (+ addition).
' - OAuth2 / Token Handling: INI_OAuth2 flag and related client mail, scopes, endpoint,
'   access token expiry; secondary equivalents; TokenExpiry (+_2).
' - Search / iSearch Feature: INI_ISearch flags, approval, URL, response masks,
'   name, tries, results, max depth, timeout, search term prompt, apply prompts
'   (markup and standard).
' - Library / Prompt Library: INI_Lib flags, file, timeout, find/apply SP (+ markup),
'   INI_PromptLib flags, paths (primary/local/transcript), PromptLibrary list,
'   PromptTitles list.
' - File / Path Configuration: Multiple INI_*Path properties (OpenSSL, HelpMeInky,
'   RedactionInstructions, SpeechModel, TTS endpoint, MyStyle, AlternateModel,
'   SpecialService, FindClause, WebAgent, DocCheck, PromptLib, FindClause local variants).
' - Misc Feature Flags & Parameters: INI_DoubleS, INI_Clean, INI_NoDash, INI_DefaultPrefix,
'   INI_Pre/PostCorrection, INI_APIEncrypted (+_2), INI_APIKeyPrefix (+_2), INI_SecondAPI,
'   INI_APIDebug, INI_UsageRestrictions, INI_Language1/2, INI_ShortcutsWordExcel,
'   INI_PromptLib, MenusAdded, INI_Model_Parameter1–4, INI_ContextMenu, INI_UpdateCheckInterval,
'   INI_UpdatePath, INI_HelpMeInkyPath, INI_RedactionInstructionsPath(+Local), Codebasis,
'   GPTSetupError, INIloaded, RDV, InitialConfigFailed (read‑only constants/flags).
' - Override Helpers: Three `Override` function overloads (String, Boolean, Integer)
'   providing controlled substitution of original values only when a supplied string
'   is interpretable (Boolean accepts localized true/false tokens; Integer uses invariant parse).
' - Responsibility: No validation, persistence, or side‑effects inside properties;
'   all lifecycle / loading handled externally via SharedLibrary configuration routines.
' =============================================================================


Option Strict Off
Option Explicit On

Imports System.Globalization
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedContext

Partial Public Class ThisAddIn

    Public Shared _context As ISharedContext = New SharedContext()

    Public Shared Property INI_APIKey As String
        Get
            Return _context.INI_APIKey
        End Get
        Set(value As String)
            _context.INI_APIKey = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack As String
        Get
            Return _context.INI_APIKeyBack
        End Get
        Set(value As String)
            _context.INI_APIKeyBack = value
        End Set
    End Property

    Public Shared Property INI_Temperature As String
        Get
            Return _context.INI_Temperature
        End Get
        Set(value As String)
            _context.INI_Temperature = value
        End Set
    End Property

    Public Shared Property INI_Timeout As Long
        Get
            Return _context.INI_Timeout
        End Get
        Set(value As Long)
            _context.INI_Timeout = value
        End Set
    End Property

    Public Shared Property INI_MaxOutputToken As Integer
        Get
            Return _context.INI_MaxOutputToken
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken = value
        End Set
    End Property
    Public Shared Property INI_Model As String
        Get
            Return _context.INI_Model
        End Get
        Set(value As String)
            _context.INI_Model = value
        End Set
    End Property

    Public Shared Property INI_Endpoint As String
        Get
            Return _context.INI_Endpoint
        End Get
        Set(value As String)
            _context.INI_Endpoint = value
        End Set
    End Property

    Public Shared Property INI_HeaderA As String
        Get
            Return _context.INI_HeaderA
        End Get
        Set(value As String)
            _context.INI_HeaderA = value
        End Set
    End Property

    Public Shared Property INI_HeaderB As String
        Get
            Return _context.INI_HeaderB
        End Get
        Set(value As String)
            _context.INI_HeaderB = value
        End Set
    End Property

    Public Shared Property INI_APICall As String
        Get
            Return _context.INI_APICall
        End Get
        Set(value As String)
            _context.INI_APICall = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object As String
        Get
            Return _context.INI_APICall_Object
        End Get
        Set(value As String)
            _context.INI_APICall_Object = value
        End Set
    End Property


    Public Shared Property INI_Response As String
        Get
            Return _context.INI_Response
        End Get
        Set(value As String)
            _context.INI_Response = value
        End Set
    End Property

    Public Shared Property INI_Anon As String
        Get
            Return _context.INI_Anon
        End Get
        Set(value As String)
            _context.INI_Anon = value
        End Set
    End Property

    Public Shared Property INI_TokenCount As String
        Get
            Return _context.INI_TokenCount
        End Get
        Set(value As String)
            _context.INI_TokenCount = value
        End Set
    End Property

    Public Shared Property INI_TokenCount_2 As String
        Get
            Return _context.INI_TokenCount_2
        End Get
        Set(value As String)
            _context.INI_TokenCount_2 = value
        End Set
    End Property

    Public Shared Property INI_DoubleS As Boolean
        Get
            Return _context.INI_DoubleS
        End Get
        Set(value As Boolean)
            _context.INI_DoubleS = value
        End Set
    End Property

    Public Shared Property INI_Clean As Boolean
        Get
            Return _context.INI_Clean
        End Get
        Set(value As Boolean)
            _context.INI_Clean = value
        End Set
    End Property

    Public Shared Property INI_Ignore As Boolean
        Get
            Return _context.INI_Ignore
        End Get
        Set(value As Boolean)
            _context.INI_Ignore = value
        End Set
    End Property

    Public Shared Property INI_Location As String
        Get
            Return _context.INI_Location
        End Get
        Set(value As String)
            _context.INI_Location = value
        End Set
    End Property
    Public Shared Property INI_NoDash As Boolean
        Get
            Return _context.INI_NoDash
        End Get
        Set(value As Boolean)
            _context.INI_NoDash = value
        End Set
    End Property

    Public Shared Property INI_DefaultPrefix As String
        Get
            Return _context.INI_DefaultPrefix
        End Get
        Set(value As String)
            _context.INI_DefaultPrefix = value
        End Set
    End Property

    Public Shared Property INI_MarkdownBubbles As Boolean
        Get
            Return _context.INI_MarkdownBubbles
        End Get
        Set(value As Boolean)
            _context.INI_MarkdownBubbles = value
        End Set
    End Property
    Public Shared Property INI_PreCorrection As String
        Get
            Return _context.INI_PreCorrection
        End Get
        Set(value As String)
            _context.INI_PreCorrection = value
        End Set
    End Property

    Public Shared Property INI_PostCorrection As String
        Get
            Return _context.INI_PostCorrection
        End Get
        Set(value As String)
            _context.INI_PostCorrection = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted As Boolean
        Get
            Return _context.INI_APIEncrypted
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix As String
        Get
            Return _context.INI_APIKeyPrefix
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodOutlook As Integer
        Get
            Return _context.INI_MarkupMethodOutlook
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodOutlook = value
        End Set
    End Property

    Public Shared Property INI_MarkupDiffCap As Integer
        Get
            Return _context.INI_MarkupDiffCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupDiffCap = value
        End Set
    End Property

    Public Shared Property INI_MarkupRegexCap As Integer
        Get
            Return _context.INI_MarkupRegexCap
        End Get
        Set(value As Integer)
            _context.INI_MarkupRegexCap = value
        End Set
    End Property

    Public Shared Property INI_OpenSSLPath As String
        Get
            Return _context.INI_OpenSSLPath
        End Get
        Set(value As String)
            _context.INI_OpenSSLPath = value
        End Set
    End Property

    Public Shared Property INI_OAuth2 As Boolean
        Get
            Return _context.INI_OAuth2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail As String
        Get
            Return _context.INI_OAuth2ClientMail
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes As String
        Get
            Return _context.INI_OAuth2Scopes
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint As String
        Get
            Return _context.INI_OAuth2Endpoint
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry As Long
        Get
            Return _context.INI_OAuth2ATExpiry
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry = value
        End Set
    End Property

    Public Shared Property INI_SecondAPI As Boolean
        Get
            Return _context.INI_SecondAPI
        End Get
        Set(value As Boolean)
            _context.INI_SecondAPI = value
        End Set
    End Property

    Public Shared Property INI_APIKey_2 As String
        Get
            Return _context.INI_APIKey_2
        End Get
        Set(value As String)
            _context.INI_APIKey_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyBack_2 As String
        Get
            Return _context.INI_APIKeyBack_2
        End Get
        Set(value As String)
            _context.INI_APIKeyBack_2 = value
        End Set
    End Property

    Public Shared Property INI_Temperature_2 As String
        Get
            Return _context.INI_Temperature_2
        End Get
        Set(value As String)
            _context.INI_Temperature_2 = value
        End Set
    End Property

    Public Shared Property INI_Timeout_2 As Long
        Get
            Return _context.INI_Timeout_2
        End Get
        Set(value As Long)
            _context.INI_Timeout_2 = value
        End Set
    End Property
    Public Shared Property INI_MaxOutputToken_2 As Integer
        Get
            Return _context.INI_MaxOutputToken_2
        End Get
        Set(value As Integer)
            _context.INI_MaxOutputToken_2 = value
        End Set
    End Property
    Public Shared Property INI_Model_2 As String
        Get
            Return _context.INI_Model_2
        End Get
        Set(value As String)
            _context.INI_Model_2 = value
        End Set
    End Property

    Public Shared Property INI_Endpoint_2 As String
        Get
            Return _context.INI_Endpoint_2
        End Get
        Set(value As String)
            _context.INI_Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderA_2 As String
        Get
            Return _context.INI_HeaderA_2
        End Get
        Set(value As String)
            _context.INI_HeaderA_2 = value
        End Set
    End Property

    Public Shared Property INI_HeaderB_2 As String
        Get
            Return _context.INI_HeaderB_2
        End Get
        Set(value As String)
            _context.INI_HeaderB_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_2 As String
        Get
            Return _context.INI_APICall_2
        End Get
        Set(value As String)
            _context.INI_APICall_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_Object_2 As String
        Get
            Return _context.INI_APICall_Object_2
        End Get
        Set(value As String)
            _context.INI_APICall_Object_2 = value
        End Set
    End Property


    Public Shared Property INI_Response_2 As String
        Get
            Return _context.INI_Response_2
        End Get
        Set(value As String)
            _context.INI_Response_2 = value
        End Set
    End Property

    Public Shared Property INI_Anon_2 As String
        Get
            Return _context.INI_Anon_2
        End Get
        Set(value As String)
            _context.INI_Anon_2 = value
        End Set
    End Property

    Public Shared Property INI_APIEncrypted_2 As Boolean
        Get
            Return _context.INI_APIEncrypted_2
        End Get
        Set(value As Boolean)
            _context.INI_APIEncrypted_2 = value
        End Set
    End Property

    Public Shared Property INI_APIKeyPrefix_2 As String
        Get
            Return _context.INI_APIKeyPrefix_2
        End Get
        Set(value As String)
            _context.INI_APIKeyPrefix_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2_2 As Boolean
        Get
            Return _context.INI_OAuth2_2
        End Get
        Set(value As Boolean)
            _context.INI_OAuth2_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ClientMail_2 As String
        Get
            Return _context.INI_OAuth2ClientMail_2
        End Get
        Set(value As String)
            _context.INI_OAuth2ClientMail_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Scopes_2 As String
        Get
            Return _context.INI_OAuth2Scopes_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Scopes_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2Endpoint_2 As String
        Get
            Return _context.INI_OAuth2Endpoint_2
        End Get
        Set(value As String)
            _context.INI_OAuth2Endpoint_2 = value
        End Set
    End Property

    Public Shared Property INI_OAuth2ATExpiry_2 As Long
        Get
            Return _context.INI_OAuth2ATExpiry_2
        End Get
        Set(value As Long)
            _context.INI_OAuth2ATExpiry_2 = value
        End Set
    End Property

    Public Shared Property INI_APIDebug As Boolean
        Get
            Return _context.INI_APIDebug
        End Get
        Set(value As Boolean)
            _context.INI_APIDebug = value
        End Set
    End Property

    Public Shared Property INI_UsageRestrictions As String
        Get
            Return _context.INI_UsageRestrictions
        End Get
        Set(value As String)
            _context.INI_UsageRestrictions = value
        End Set
    End Property

    Public Shared Property INI_Language1 As String
        Get
            Return _context.INI_Language1
        End Get
        Set(value As String)
            _context.INI_Language1 = value
        End Set
    End Property

    Public Shared Property INI_Language2 As String
        Get
            Return _context.INI_Language2
        End Get
        Set(value As String)
            _context.INI_Language2 = value
        End Set
    End Property

    Public Shared Property INI_KeepFormat1 As Boolean
        Get
            Return _context.INI_KeepFormat1
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat1 = value
        End Set
    End Property

    Public Shared Property INI_MarkdownConvert As Boolean
        Get
            Return _context.INI_MarkdownConvert
        End Get
        Set(value As Boolean)
            _context.INI_MarkdownConvert = value
        End Set
    End Property

    Public Shared Property INI_KeepFormat2 As Boolean
        Get
            Return _context.INI_KeepFormat2
        End Get
        Set(value As Boolean)
            _context.INI_KeepFormat2 = value
        End Set
    End Property
    Public Shared Property INI_KeepParaFormatInline As Boolean
        Get
            Return _context.INI_KeepParaFormatInline
        End Get
        Set(value As Boolean)
            _context.INI_KeepParaFormatInline = value
        End Set
    End Property

    Public Shared Property INI_KeepFormatCap As Integer
        Get
            Return _context.INI_KeepFormatCap
        End Get
        Set(value As Integer)
            _context.INI_KeepFormatCap = value
        End Set
    End Property
    Public Shared Property INI_ReplaceText1 As Boolean
        Get
            Return _context.INI_ReplaceText1
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText1 = value
        End Set
    End Property

    Public Shared Property INI_ReplaceText2 As Boolean
        Get
            Return _context.INI_ReplaceText2
        End Get
        Set(value As Boolean)
            _context.INI_ReplaceText2 = value
        End Set
    End Property

    Public Shared Property INI_ReplaceText2Override As String
        Get
            Return _context.INI_ReplaceText2Override
        End Get
        Set(value As String)
            _context.INI_ReplaceText2Override = value
        End Set
    End Property


    Public Shared Property INI_DoMarkupOutlook As Boolean
        Get
            Return _context.INI_DoMarkupOutlook
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupOutlook = value
        End Set
    End Property

    Public Shared Property INI_DoMarkupWord As Boolean
        Get
            Return _context.INI_DoMarkupWord
        End Get
        Set(value As Boolean)
            _context.INI_DoMarkupWord = value
        End Set
    End Property

    Public Shared Property SP_Translate As String
        Get
            Return _context.SP_Translate
        End Get
        Set(value As String)
            _context.SP_Translate = value
        End Set
    End Property

    Public Shared Property SP_Correct As String
        Get
            Return _context.SP_Correct
        End Get
        Set(value As String)
            _context.SP_Correct = value
        End Set
    End Property

    Public Shared Property SP_Improve As String
        Get
            Return _context.SP_Improve
        End Get
        Set(value As String)
            _context.SP_Improve = value
        End Set
    End Property

    Public Shared Property SP_Explain As String
        Get
            Return _context.SP_Explain
        End Get
        Set(value As String)
            _context.SP_Explain = value
        End Set
    End Property

    Public Shared Property SP_FindClause As String
        Get
            Return _context.SP_FindClause
        End Get
        Set(value As String)
            _context.SP_FindClause = value
        End Set
    End Property

    Public Shared Property SP_FindClause_Clean As String
        Get
            Return _context.SP_FindClause_Clean
        End Get
        Set(value As String)
            _context.SP_FindClause_Clean = value
        End Set
    End Property

    Public Shared Property SP_ApplyDocStyle As String
        Get
            Return _context.SP_ApplyDocStyle
        End Get
        Set(value As String)
            _context.SP_ApplyDocStyle = value
        End Set
    End Property

    Public Shared Property SP_ApplyDocStyle_NumberingHint As String
        Get
            Return _context.SP_ApplyDocStyle_NumberingHint
        End Get
        Set(value As String)
            _context.SP_ApplyDocStyle_NumberingHint = value
        End Set
    End Property


    Public Shared Property SP_DocCheck_Clause As String
        Get
            Return _context.SP_DocCheck_Clause
        End Get
        Set(value As String)
            _context.SP_DocCheck_Clause = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClause As String
        Get
            Return _context.SP_DocCheck_MultiClause
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClause = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClauseSum As String
        Get
            Return _context.SP_DocCheck_MultiClauseSum
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClauseSum = value
        End Set
    End Property

    Public Shared Property SP_DocCheck_MultiClauseSum_Bubbles As String
        Get
            Return _context.SP_DocCheck_MultiClauseSum_Bubbles
        End Get
        Set(value As String)
            _context.SP_DocCheck_MultiClauseSum_Bubbles = value
        End Set
    End Property

    Public Shared Property SP_SuggestTitles As String
        Get
            Return _context.SP_SuggestTitles
        End Get
        Set(value As String)
            _context.SP_SuggestTitles = value
        End Set
    End Property

    Public Shared Property SP_Friendly As String
        Get
            Return _context.SP_Friendly
        End Get
        Set(value As String)
            _context.SP_Friendly = value
        End Set
    End Property

    Public Shared Property SP_Convincing As String
        Get
            Return _context.SP_Convincing
        End Get
        Set(value As String)
            _context.SP_Convincing = value
        End Set
    End Property

    Public Shared Property SP_NoFillers As String
        Get
            Return _context.SP_NoFillers
        End Get
        Set(value As String)
            _context.SP_NoFillers = value
        End Set
    End Property

    Public Shared Property SP_Podcast As String
        Get
            Return _context.SP_Podcast
        End Get
        Set(value As String)
            _context.SP_Podcast = value
        End Set
    End Property


    Public Shared Property SP_MyStyle_Word As String
        Get
            Return _context.SP_MyStyle_Word
        End Get
        Set(value As String)
            _context.SP_MyStyle_Word = value
        End Set
    End Property

    Public Shared Property SP_MyStyle_Outlook As String
        Get
            Return _context.SP_MyStyle_Outlook
        End Get
        Set(value As String)
            _context.SP_MyStyle_Outlook = value
        End Set
    End Property

    Public Shared Property SP_MyStyle_Apply As String
        Get
            Return _context.SP_MyStyle_Apply
        End Get
        Set(value As String)
            _context.SP_MyStyle_Apply = value
        End Set
    End Property


    Public Shared Property SP_Shorten As String
        Get
            Return _context.SP_Shorten
        End Get
        Set(value As String)
            _context.SP_Shorten = value
        End Set
    End Property

    Public Shared Property SP_Filibuster As String
        Get
            Return _context.SP_Filibuster
        End Get
        Set(value As String)
            _context.SP_Filibuster = value
        End Set
    End Property

    Public Shared Property SP_ArgueAgainst As String
        Get
            Return _context.SP_ArgueAgainst
        End Get
        Set(value As String)
            _context.SP_ArgueAgainst = value
        End Set
    End Property

    Public Shared Property SP_InsertClipboard As String
        Get
            Return _context.SP_InsertClipboard
        End Get
        Set(value As String)
            _context.SP_InsertClipboard = value
        End Set
    End Property


    Public Shared Property SP_Summarize As String
        Get
            Return _context.SP_Summarize
        End Get
        Set(value As String)
            _context.SP_Summarize = value
        End Set
    End Property

    Public Shared Property SP_Markup As String
        Get
            Return _context.SP_Markup
        End Get
        Set(value As String)
            _context.SP_Markup = value
        End Set
    End Property

    Public Shared Property SP_MailReply As String
        Get
            Return _context.SP_MailReply
        End Get
        Set(value As String)
            _context.SP_MailReply = value
        End Set
    End Property

    Public Shared Property SP_MailSumup As String
        Get
            Return _context.SP_MailSumup
        End Get
        Set(value As String)
            _context.SP_MailSumup = value
        End Set
    End Property

    Public Shared Property SP_MailSumup2 As String
        Get
            Return _context.SP_MailSumup2
        End Get
        Set(value As String)
            _context.SP_MailSumup2 = value
        End Set
    End Property

    Public Shared Property SP_FreestyleText As String
        Get
            Return _context.SP_FreestyleText
        End Get
        Set(value As String)
            _context.SP_FreestyleText = value
        End Set
    End Property

    Public Shared Property SP_FreestyleNoText As String
        Get
            Return _context.SP_FreestyleNoText
        End Get
        Set(value As String)
            _context.SP_FreestyleNoText = value
        End Set
    End Property

    Public Shared Property SP_SwitchParty As String
        Get
            Return _context.SP_SwitchParty
        End Get
        Set(value As String)
            _context.SP_SwitchParty = value
        End Set
    End Property

    Public Shared Property SP_Anonymize As String
        Get
            Return _context.SP_Anonymize
        End Get
        Set(value As String)
            _context.SP_Anonymize = value
        End Set
    End Property


    Public Shared Property SP_Extract As String
        Get
            Return _context.SP_Extract
        End Get
        Set(value As String)
            _context.SP_Extract = value
        End Set
    End Property

    Public Shared Property SP_ExtractSchema As String
        Get
            Return _context.SP_ExtractSchema
        End Get
        Set(value As String)
            _context.SP_ExtractSchema = value
        End Set
    End Property

    Public Shared Property SP_MergeDateRows As String
        Get
            Return _context.SP_MergeDateRows
        End Get
        Set(value As String)
            _context.SP_MergeDateRows = value
        End Set
    End Property

    Public Shared Property SP_Redact As String
        Get
            Return _context.SP_Redact
        End Get
        Set(value As String)
            _context.SP_Redact = value
        End Set
    End Property

    Public Shared Property SP_Rename As String
        Get
            Return _context.SP_Rename
        End Get
        Set(value As String)
            _context.SP_Rename = value
        End Set
    End Property


    Public Shared Property SP_CheckforII As String
        Get
            Return _context.SP_CheckforII
        End Get
        Set(value As String)
            _context.SP_CheckforII = value
        End Set
    End Property



    Public Shared Property SP_ContextSearch As String
        Get
            Return _context.SP_ContextSearch
        End Get
        Set(value As String)
            _context.SP_ContextSearch = value
        End Set
    End Property

    Public Shared Property SP_ContextSearchMulti As String
        Get
            Return _context.SP_ContextSearchMulti
        End Get
        Set(value As String)
            _context.SP_ContextSearchMulti = value
        End Set
    End Property

    Public Shared Property SP_RangeOfCells As String
        Get
            Return _context.SP_RangeOfCells
        End Get
        Set(value As String)
            _context.SP_RangeOfCells = value
        End Set
    End Property

    Public Shared Property SP_ParseFile As String
        Get
            Return _context.SP_ParseFile
        End Get
        Set(value As String)
            _context.SP_ParseFile = value
        End Set
    End Property

    Public Shared Property SP_Ignore As String
        Get
            Return _context.SP_Ignore
        End Get
        Set(value As String)
            _context.SP_Ignore = value
        End Set
    End Property

    Public Shared Property SP_WriteNeatly As String
        Get
            Return _context.SP_WriteNeatly
        End Get
        Set(value As String)
            _context.SP_WriteNeatly = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepFormulasIntact As String
        Get
            Return _context.SP_Add_KeepFormulasIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepFormulasIntact = value
        End Set
    End Property
    Public Shared Property SP_Add_KeepHTMLIntact As String
        Get
            Return _context.SP_Add_KeepHTMLIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepHTMLIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_KeepInlineIntact As String
        Get
            Return _context.SP_Add_KeepInlineIntact
        End Get
        Set(value As String)
            _context.SP_Add_KeepInlineIntact = value
        End Set
    End Property

    Public Shared Property SP_Add_Bubbles As String
        Get
            Return _context.SP_Add_Bubbles
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles = value
        End Set
    End Property

    Public Shared Property SP_Add_BubblesReply As String
        Get
            Return _context.SP_Add_BubblesReply
        End Get
        Set(value As String)
            _context.SP_Add_BubblesReply = value
        End Set
    End Property

    Public Shared Property SP_Add_BubblesExtract As String
        Get
            Return _context.SP_Add_BubblesExtract
        End Get
        Set(value As String)
            _context.SP_Add_BubblesExtract = value
        End Set
    End Property


    Public Shared Property SP_Add_Bubbles_Format As String
        Get
            Return _context.SP_Add_Bubbles_Format
        End Get
        Set(value As String)
            _context.SP_Add_Bubbles_Format = value
        End Set
    End Property


    Public Shared Property SP_Add_Batch As String
        Get
            Return _context.SP_Add_Batch
        End Get
        Set(value As String)
            _context.SP_Add_Batch = value
        End Set
    End Property

    Public Shared Property SP_Add_Tooling As String
        Get
            Return _context.SP_Add_Tooling
        End Get
        Set(value As String)
            _context.SP_Add_Tooling = value
        End Set
    End Property


    Public Shared Property SP_BubblesExcel As String
        Get
            Return _context.SP_BubblesExcel
        End Get
        Set(value As String)
            _context.SP_BubblesExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_Revisions As String
        Get
            Return _context.SP_Add_Revisions
        End Get
        Set(value As String)
            _context.SP_Add_Revisions = value
        End Set
    End Property

    Public Shared Property SP_MarkupRegex As String
        Get
            Return _context.SP_MarkupRegex
        End Get
        Set(value As String)
            _context.SP_MarkupRegex = value
        End Set
    End Property


    Public Shared Property SP_ChatWord As String
        Get
            Return _context.SP_ChatWord
        End Get
        Set(value As String)
            _context.SP_ChatWord = value
        End Set
    End Property

    Public Shared Property SP_Chat As String
        Get
            Return _context.SP_Chat
        End Get
        Set(value As String)
            _context.SP_Chat = value
        End Set
    End Property

    Public Shared Property SP_HelpMe As String
        Get
            Return _context.SP_HelpMe
        End Get
        Set(value As String)
            _context.SP_HelpMe = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatWord_Commands As String
        Get
            Return _context.SP_Add_ChatWord_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatWord_Commands = value
        End Set
    End Property

    Public Shared Property SP_Add_Chat_NoCommands As String
        Get
            Return _context.SP_Add_Chat_NoCommands
        End Get
        Set(value As String)
            _context.SP_Add_Chat_NoCommands = value
        End Set
    End Property


    Public Shared Property SP_ChatExcel As String
        Get
            Return _context.SP_ChatExcel
        End Get
        Set(value As String)
            _context.SP_ChatExcel = value
        End Set
    End Property

    Public Shared Property SP_Add_ChatExcel_Commands As String
        Get
            Return _context.SP_Add_ChatExcel_Commands
        End Get
        Set(value As String)
            _context.SP_Add_ChatExcel_Commands = value
        End Set
    End Property

    Public Shared Property INI_ChatCap As Integer
        Get
            Return _context.INI_ChatCap
        End Get
        Set(value As Integer)
            _context.INI_ChatCap = value
        End Set
    End Property

    Public Shared ReadOnly Property RDV As String = "Outlook (" & Version & ")"
    Public Shared ReadOnly Property InitialConfigFailed As Boolean = False
    Public Shared Property DecodedAPI As String
        Get
            Return _context.DecodedAPI
        End Get
        Set(value As String)
            _context.DecodedAPI = value
        End Set
    End Property

    Public Shared Property DecodedAPI_2 As String
        Get
            Return _context.DecodedAPI_2
        End Get
        Set(value As String)
            _context.DecodedAPI_2 = value
        End Set
    End Property

    Public Shared Property TokenExpiry As DateTime
        Get
            Return _context.TokenExpiry
        End Get
        Set(value As DateTime)
            _context.TokenExpiry = value
        End Set
    End Property

    Public Shared Property TokenExpiry_2 As DateTime
        Get
            Return _context.TokenExpiry_2
        End Get
        Set(value As DateTime)
            _context.TokenExpiry_2 = value
        End Set
    End Property

    Public Shared Property Codebasis As String
        Get
            Return _context.Codebasis
        End Get
        Set(value As String)
            _context.Codebasis = value
        End Set
    End Property

    Public Shared Property GPTSetupError As Boolean
        Get
            Return _context.GPTSetupError
        End Get
        Set(value As Boolean)
            _context.GPTSetupError = value
        End Set
    End Property

    Public Shared Property INIloaded As Boolean
        Get
            Return _context.INIloaded
        End Get
        Set(value As Boolean)
            _context.INIloaded = value
        End Set
    End Property



    Public Shared Property INI_ISearch As Boolean
        Get
            Return _context.INI_ISearch
        End Get
        Set(value As Boolean)
            _context.INI_ISearch = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Approve As Boolean
        Get
            Return _context.INI_ISearch_Approve
        End Get
        Set(value As Boolean)
            _context.INI_ISearch_Approve = value
        End Set
    End Property

    Public Shared Property INI_ISearch_URL As String
        Get
            Return _context.INI_ISearch_URL
        End Get
        Set(value As String)
            _context.INI_ISearch_URL = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseURLStart As String
        Get
            Return _context.INI_ISearch_ResponseURLStart
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseURLStart = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask1 As String
        Get
            Return _context.INI_ISearch_ResponseMask1
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask1 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_ResponseMask2 As String
        Get
            Return _context.INI_ISearch_ResponseMask2
        End Get
        Set(value As String)
            _context.INI_ISearch_ResponseMask2 = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Name As String
        Get
            Return _context.INI_ISearch_Name
        End Get
        Set(value As String)
            _context.INI_ISearch_Name = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Tries As Integer
        Get
            Return _context.INI_ISearch_Tries
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Tries = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Results As Integer
        Get
            Return _context.INI_ISearch_Results
        End Get
        Set(value As Integer)
            _context.INI_ISearch_Results = value
        End Set
    End Property

    Public Shared Property INI_ISearch_MaxDepth As Integer
        Get
            Return _context.INI_ISearch_MaxDepth
        End Get
        Set(value As Integer)
            _context.INI_ISearch_MaxDepth = value
        End Set
    End Property

    Public Shared Property INI_ISearch_Timeout As Long
        Get
            Return _context.INI_ISearch_Timeout
        End Get
        Set(value As Long)
            _context.INI_ISearch_Timeout = value
        End Set
    End Property

    Public Shared Property INI_ISearch_SearchTerm_SP As String
        Get
            Return _context.INI_ISearch_SearchTerm_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_SearchTerm_SP = value
        End Set
    End Property


    Public Shared Property INI_ISearch_Apply_SP As String
        Get
            Return _context.INI_ISearch_Apply_SP
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP = value
        End Set
    End Property
    Public Shared Property INI_ISearch_Apply_SP_Markup As String
        Get
            Return _context.INI_ISearch_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_ISearch_Apply_SP_Markup = value
        End Set
    End Property
    Public Shared Property INI_Lib As Boolean
        Get
            Return _context.INI_Lib
        End Get
        Set(value As Boolean)
            _context.INI_Lib = value
        End Set
    End Property

    Public Shared Property INI_Lib_File As String
        Get
            Return _context.INI_Lib_File
        End Get
        Set(value As String)
            _context.INI_Lib_File = value
        End Set
    End Property

    Public Shared Property INI_Lib_Timeout As Long
        Get
            Return _context.INI_Lib_Timeout
        End Get
        Set(value As Long)
            _context.INI_Lib_Timeout = value
        End Set
    End Property

    Public Shared Property INI_Lib_Find_SP As String
        Get
            Return _context.INI_Lib_Find_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Find_SP = value
        End Set
    End Property


    Public Shared Property INI_Lib_Apply_SP As String
        Get
            Return _context.INI_Lib_Apply_SP
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP = value
        End Set
    End Property

    Public Shared Property INI_Lib_Apply_SP_Markup As String
        Get
            Return _context.INI_Lib_Apply_SP_Markup
        End Get
        Set(value As String)
            _context.INI_Lib_Apply_SP_Markup = value
        End Set
    End Property




    Public Shared Property INI_MarkupMethodHelper As Integer
        Get
            Return _context.INI_MarkupMethodHelper
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodHelper = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodWord As Integer
        Get
            Return _context.INI_MarkupMethodWord
        End Get
        Set(value As Integer)
            _context.INI_MarkupMethodWord = value
        End Set
    End Property

    Public Shared Property INI_MarkupMethodWordOverride As String
        Get
            Return _context.INI_MarkupMethodWordOverride
        End Get
        Set(value As String)
            _context.INI_MarkupMethodWordOverride = value
        End Set
    End Property


    Public Shared Property INI_MarkupMethodOutlookOverride As String
        Get
            Return _context.INI_MarkupMethodOutlookOverride
        End Get
        Set(value As String)
            _context.INI_MarkupMethodOutlookOverride = value
        End Set
    End Property


    Public Shared Property INI_ContextMenu As Boolean
        Get
            Return _context.INI_ContextMenu
        End Get
        Set(value As Boolean)
            _context.INI_ContextMenu = value
        End Set
    End Property

    Public Shared Property INI_UpdateCheckInterval As Integer
        Get
            Return _context.INI_UpdateCheckInterval
        End Get
        Set(value As Integer)
            _context.INI_UpdateCheckInterval = value
        End Set
    End Property

    Public Shared Property INI_UpdatePath As String
        Get
            Return _context.INI_UpdatePath
        End Get
        Set(value As String)
            _context.INI_UpdatePath = value
        End Set
    End Property

    Public Shared Property INI_HelpMeInkyPath As String
        Get
            Return _context.INI_HelpMeInkyPath
        End Get
        Set(value As String)
            _context.INI_HelpMeInkyPath = value
        End Set
    End Property

    Public Shared Property INI_DiscussInkyPath As String
        Get
            Return _context.INI_DiscussInkyPath
        End Get
        Set(value As String)
            _context.INI_DiscussInkyPath = value
        End Set
    End Property

    Public Shared Property INI_DiscussInkyPathLocal As String
        Get
            Return _context.INI_DiscussInkyPathLocal
        End Get
        Set(value As String)
            _context.INI_DiscussInkyPathLocal = value
        End Set
    End Property
    Public Shared Property INI_RedactionInstructionsPath As String
        Get
            Return _context.INI_RedactionInstructionsPath
        End Get
        Set(value As String)
            _context.INI_RedactionInstructionsPath = value
        End Set
    End Property

    Public Shared Property INI_RedactionInstructionsPathLocal As String
        Get
            Return _context.INI_RedactionInstructionsPathLocal
        End Get
        Set(value As String)
            _context.INI_RedactionInstructionsPathLocal = value
        End Set
    End Property

    Public Shared Property INI_ExtractorPath As String
        Get
            Return _context.INI_ExtractorPath
        End Get
        Set(value As String)
            _context.INI_ExtractorPath = value
        End Set
    End Property

    Public Shared Property INI_ExtractorPathLocal As String
        Get
            Return _context.INI_ExtractorPathLocal
        End Get
        Set(value As String)
            _context.INI_ExtractorPathLocal = value
        End Set
    End Property

    Public Shared Property INI_RenameLibPath As String
        Get
            Return _context.INI_RenameLibPath
        End Get
        Set(value As String)
            _context.INI_RenameLibPath = value
        End Set
    End Property

    Public Shared Property INI_RenameLibPathLocal As String
        Get
            Return _context.INI_RenameLibPathLocal
        End Get
        Set(value As String)
            _context.INI_RenameLibPathLocal = value
        End Set
    End Property

    Public Shared Property INI_SpeechModelPath As String
        Get
            Return _context.INI_SpeechModelPath
        End Get
        Set(value As String)
            _context.INI_SpeechModelPath = value
        End Set
    End Property

    Public Shared Property INI_TTSEndpoint As String
        Get
            Return _context.INI_TTSEndpoint
        End Get
        Set(value As String)
            _context.INI_TTSEndpoint = value
        End Set
    End Property

    Public Shared Property INI_ShortcutsWordExcel As String
        Get
            Return _context.INI_ShortcutsWordExcel
        End Get
        Set(value As String)
            _context.INI_ShortcutsWordExcel = value
        End Set
    End Property

    Public Shared Property INI_PromptLib As Boolean
        Get
            Return _context.INI_PromptLib
        End Get
        Set(value As Boolean)
            _context.INI_PromptLib = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPath As String
        Get
            Return _context.INI_PromptLibPath
        End Get
        Set(value As String)
            _context.INI_PromptLibPath = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPathLocal As String
        Get
            Return _context.INI_PromptLibPathLocal
        End Get
        Set(value As String)
            _context.INI_PromptLibPathLocal = value
        End Set
    End Property


    Public Shared Property INI_MyStylePath As String
        Get
            Return _context.INI_MyStylePath
        End Get
        Set(value As String)
            _context.INI_MyStylePath = value
        End Set
    End Property

    Public Shared Property INI_AlternateModelPath As String
        Get
            Return _context.INI_AlternateModelPath
        End Get
        Set(value As String)
            _context.INI_AlternateModelPath = value
        End Set
    End Property

    Public Shared Property INI_SpecialServicePath As String
        Get
            Return _context.INI_SpecialServicePath
        End Get
        Set(value As String)
            _context.INI_SpecialServicePath = value
        End Set
    End Property

    Public Shared Property INI_FindClausePath As String
        Get
            Return _context.INI_FindClausePath
        End Get
        Set(value As String)
            _context.INI_FindClausePath = value
        End Set
    End Property

    Public Shared Property INI_FindClausePathLocal As String
        Get
            Return _context.INI_FindClausePathLocal
        End Get
        Set(value As String)
            _context.INI_FindClausePathLocal = value
        End Set
    End Property

    Public Shared Property INI_WebAgentPath As String
        Get
            Return _context.INI_WebAgentPath
        End Get
        Set(value As String)
            _context.INI_WebAgentPath = value
        End Set
    End Property

    Public Shared Property INI_WebAgentPathLocal As String
        Get
            Return _context.INI_WebAgentPathLocal
        End Get
        Set(value As String)
            _context.INI_WebAgentPathLocal = value
        End Set
    End Property

    Public Shared Property INI_DocCheckPath As String
        Get
            Return _context.INI_DocCheckPath
        End Get
        Set(value As String)
            _context.INI_DocCheckPath = value
        End Set
    End Property

    Public Shared Property INI_DocCheckPathLocal As String
        Get
            Return _context.INI_DocCheckPathLocal
        End Get
        Set(value As String)
            _context.INI_DocCheckPathLocal = value
        End Set
    End Property

    Public Shared Property INI_DocStylePath As String
        Get
            Return _context.INI_DocStylePath
        End Get
        Set(value As String)
            _context.INI_DocStylePath = value
        End Set
    End Property

    Public Shared Property INI_DocStylePathLocal As String
        Get
            Return _context.INI_DocStylePathLocal
        End Get
        Set(value As String)
            _context.INI_DocStylePathLocal = value
        End Set
    End Property

    Public Shared Property INI_PromptLibPath_Transcript As String
        Get
            Return _context.INI_PromptLibPath_Transcript
        End Get
        Set(value As String)
            _context.INI_PromptLibPath_Transcript = value
        End Set
    End Property

    Public Shared Property PromptLibrary() As List(Of String)
        Get
            Return _context.PromptLibrary
        End Get
        Set(value As List(Of String))
            _context.PromptLibrary = value
        End Set
    End Property

    Public Shared Property PromptTitles() As List(Of String)
        Get
            Return _context.PromptTitles
        End Get
        Set(value As List(Of String))
            _context.PromptTitles = value
        End Set
    End Property

    Public Shared Property MenusAdded As Boolean
        Get
            Return _context.MenusAdded
        End Get
        Set(value As Boolean)
            _context.MenusAdded = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter1 As String
        Get
            Return _context.INI_Model_Parameter1
        End Get
        Set(value As String)
            _context.INI_Model_Parameter1 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter2 As String
        Get
            Return _context.INI_Model_Parameter2
        End Get
        Set(value As String)
            _context.INI_Model_Parameter2 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter3 As String
        Get
            Return _context.INI_Model_Parameter3
        End Get
        Set(value As String)
            _context.INI_Model_Parameter3 = value
        End Set
    End Property

    Public Shared Property INI_Model_Parameter4 As String
        Get
            Return _context.INI_Model_Parameter4
        End Get
        Set(value As String)
            _context.INI_Model_Parameter4 = value
        End Set
    End Property

    Public Shared Property SP_FindPrompts As String
        Get
            Return _context.SP_FindPrompts
        End Get
        Set(value As String)
            _context.SP_FindPrompts = value
        End Set
    End Property

    Public Shared Property SP_MergePrompt As String
        Get
            Return _context.SP_MergePrompt
        End Get
        Set(value As String)
            _context.SP_MergePrompt = value
        End Set
    End Property
    Public Shared Property SP_Add_MergePrompt As String
        Get
            Return _context.SP_Add_MergePrompt
        End Get
        Set(value As String)
            _context.SP_Add_MergePrompt = value
        End Set
    End Property


    Public Shared Property INI_UpdateIniSilentMode As Integer
        Get
            Return _context.INI_UpdateIniSilentMode
        End Get
        Set(value As Integer)
            _context.INI_UpdateIniSilentMode = value
        End Set
    End Property

    Public Shared Property INI_UpdateIniSilentLog As Boolean
        Get
            Return _context.INI_UpdateIniSilentLog
        End Get
        Set(value As Boolean)
            _context.INI_UpdateIniSilentLog = value
        End Set
    End Property

    Public Shared Property INI_UpdateIni As Boolean
        Get
            Return _context.INI_UpdateIni
        End Get
        Set(value As Boolean)
            _context.INI_UpdateIni = value
        End Set
    End Property

    Public Shared Property INI_UpdateIniAllowRemote As Boolean
        Get
            Return _context.INI_UpdateIniAllowRemote
        End Get
        Set(value As Boolean)
            _context.INI_UpdateIniAllowRemote = value
        End Set
    End Property

    Public Shared Property INI_UpdateIniNoSignature As Boolean
        Get
            Return _context.INI_UpdateIniNoSignature
        End Get
        Set(value As Boolean)
            _context.INI_UpdateIniNoSignature = value
        End Set
    End Property

    Public Shared Property INI_UpdateSource As String
        Get
            Return _context.INI_UpdateSource
        End Get
        Set(value As String)
            _context.INI_UpdateSource = value
        End Set
    End Property

    Public Shared Property INI_UpdateIniIgnoreOverride As String
        Get
            Return _context.INI_UpdateIniIgnoreOverride
        End Get
        Set(value As String)
            _context.INI_UpdateIniIgnoreOverride = value
        End Set
    End Property

    Public Shared Property Ignore As String
        Get
            Return _context.Ignore
        End Get
        Set(value As String)
            _context.Ignore = value
        End Set
    End Property

    Public Shared Property Location As String
        Get
            Return _context.Location
        End Get
        Set(value As String)
            _context.Location = value
        End Set
    End Property

    Public Shared Property INI_NoHelperDownload As Boolean
        Get
            Return _context.INI_NoHelperDownload
        End Get
        Set(value As Boolean)
            _context.INI_NoHelperDownload = value
        End Set
    End Property

    ' Tooling / tool-call settings 

    Public Shared Property INI_ToolingLogWindow As Boolean
        Get
            Return _context.INI_ToolingLogWindow
        End Get
        Set(value As Boolean)
            _context.INI_ToolingLogWindow = value
        End Set
    End Property

    Public Shared Property INI_ToolingDryRun As Boolean
        Get
            Return _context.INI_ToolingDryRun
        End Get
        Set(value As Boolean)
            _context.INI_ToolingDryRun = value
        End Set
    End Property

    Public Shared Property INI_ToolingMaximumIterations As Integer
        Get
            Return _context.INI_ToolingMaximumIterations
        End Get
        Set(value As Integer)
            _context.INI_ToolingMaximumIterations = value
        End Set
    End Property

    Public Shared Property INI_APICall_ToolInstructions_2 As String
        Get
            Return _context.INI_APICall_ToolInstructions_2
        End Get
        Set(value As String)
            _context.INI_APICall_ToolInstructions_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_ToolInstructions_Template_2 As String
        Get
            Return _context.INI_APICall_ToolInstructions_Template_2
        End Get
        Set(value As String)
            _context.INI_APICall_ToolInstructions_Template_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_ToolResponses_2 As String
        Get
            Return _context.INI_APICall_ToolResponses_2
        End Get
        Set(value As String)
            _context.INI_APICall_ToolResponses_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_ToolResponses_Template_2 As String
        Get
            Return _context.INI_APICall_ToolResponses_Template_2
        End Get
        Set(value As String)
            _context.INI_APICall_ToolResponses_Template_2 = value
        End Set
    End Property

    Public Shared Property INI_APICall_ToolCallPart_Template_2 As String
        Get
            Return _context.INI_APICall_ToolCallPart_Template_2
        End Get
        Set(value As String)
            _context.INI_APICall_ToolCallPart_Template_2 = value
        End Set
    End Property

    Public Shared Property INI_ToolCallDetectionPattern_2 As String
        Get
            Return _context.INI_ToolCallDetectionPattern_2
        End Get
        Set(value As String)
            _context.INI_ToolCallDetectionPattern_2 = value
        End Set
    End Property

    Public Shared Property INI_ToolCallExtractionMap_2 As String
        Get
            Return _context.INI_ToolCallExtractionMap_2
        End Get
        Set(value As String)
            _context.INI_ToolCallExtractionMap_2 = value
        End Set
    End Property

    ''' <summary>
    ''' Returns the supplied <paramref name="OverrideValue"/> when it is not null, empty,
    ''' or whitespace; otherwise returns the original string value unchanged.
    ''' </summary>
    ''' <param name="Original">Current (fallback) string value.</param>
    ''' <param name="OverrideValue">Proposed replacement; trimmed only for the emptiness check.</param>
    ''' <returns><paramref name="OverrideValue"/> if it contains non-whitespace characters; otherwise <paramref name="Original"/>.</returns>
    ''' <remarks>
    ''' No additional validation or transformation is performed. Use this overload for simple
    ''' override semantics where an empty user input should keep the existing configuration value.
    ''' </remarks>
    Public Function Override(Original As String, OverrideValue As String) As String
        If String.IsNullOrWhiteSpace(OverrideValue) Then
            Return Original
        End If
        Return OverrideValue
    End Function

    ''' <summary>
    ''' Attempts to override a Boolean configuration value using a user-supplied string.
    ''' Returns the interpreted Boolean when <paramref name="OverrideValue"/> matches a known
    ''' true/false token; otherwise the original value is preserved.
    ''' </summary>
    ''' <param name="Original">Existing Boolean value to keep if no valid override is provided.</param>
    ''' <param name="OverrideValue">
    ''' A case-insensitive string token. Accepted true tokens: 1, yes, ja, wahr, true.
    ''' Accepted false tokens: 0, no, nein, falsch, false. Empty, null, or unrecognized values revert to <paramref name="Original"/>.
    ''' </param>
    ''' <returns>The overridden Boolean if recognized; otherwise <paramref name="Original"/>.</returns>
    ''' <remarks>
    ''' Includes simple localization for German (ja/wahr and nein/falsch). Tokens are trimmed before evaluation.
    ''' </remarks>
    Public Function Override(Original As Boolean, OverrideValue As String) As Boolean
        If String.IsNullOrWhiteSpace(OverrideValue) Then
            Return Original
        End If

        Dim s As String = OverrideValue.Trim().ToLowerInvariant()

        If s = "1" OrElse s = "yes" OrElse s = "ja" OrElse s = "wahr" OrElse s = "true" Then
            Return True
        End If

        If s = "0" OrElse s = "no" OrElse s = "nein" OrElse s = "falsch" OrElse s = "false" Then
            Return False
        End If

        ' Not interpretable → keep original
        Return Original
    End Function

    ''' <summary>
    ''' Attempts to override an Integer configuration value using a user-supplied string.
    ''' Parses <paramref name="OverrideValue"/> as a 32-bit integer (InvariantCulture); if parsing fails,
    ''' the original value is returned unchanged.
    ''' </summary>
    ''' <param name="Original">Existing integer value to retain if parsing is unsuccessful.</param>
    ''' <param name="OverrideValue">String representation of a 32-bit integer. Trimmed before parsing; may be null or empty.</param>
    ''' <returns>The parsed integer when valid; otherwise <paramref name="Original"/>.</returns>
    ''' <remarks>
    ''' Uses <see cref="Globalization.NumberStyles.Integer"/> with <see cref="CultureInfo.InvariantCulture"/>.
    ''' No range checking beyond Int32 parse semantics is performed.
    ''' </remarks>
    Public Function Override(Original As Integer, OverrideValue As String) As Integer
        If String.IsNullOrWhiteSpace(OverrideValue) Then
            Return Original
        End If
        Dim parsed As Integer
        If Integer.TryParse(OverrideValue.Trim(), Globalization.NumberStyles.Integer, CultureInfo.InvariantCulture, parsed) Then
            Return parsed
        End If
        Return Original
    End Function


End Class
