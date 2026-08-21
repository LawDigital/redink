' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' 21.8.2026
'
' The compiled version of Red Ink also ...
'
' Includes DiffPlex in unchanged form; Copyright (c) 2023 Matthew Manela; licensed under the Apache-2.0 license (http://www.apache.org/licenses/LICENSE-2.0) at GitHub (https://github.com/mmanela/diffplex).
' Includes Newtonsoft.Json in unchanged form; Copyright (c) 2023 James Newton-King; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.newtonsoft.com/json
' Includes HtmlAgilityPack in unchanged form; Copyright (c) 2024 ZZZ Projects, Simon Mourrier,Jeff Klawiter,Stephan Grell; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://html-agility-pack.net/
' Includes Bouncycastle.Cryptography in unchanged form; Copyright (c) 2024 Legion of the Bouncy Castle Inc.; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://www.bouncycastle.org/download/bouncy-castle-c/
' Includes PdfPig in unchanged form; Copyright (c) 2024 UglyToad, EliotJones PdfPig, BobLd; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/UglyToad/PdfPig
' Includes MarkDig in unchanged form; Copyright (c) 2024 Alexandre Mutel; licensed under the BSD 2 Clause (Simplified) license (https://licenses.nuget.org/BSD-2-Clause) at https://github.com/xoofx/markdig
' Includes NAudio and components in unchanged form; Copyright (c) 2020 Mark Heath; licensed under a proprietary open source license (https://www.nuget.org/packages/NAudio/2.2.1/license) at https://github.com/naudio/NAudio
' Includes Vosk in unchanged form; Copyright (c) 2022 Alpha Cephei Inc.; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://alphacephei.com/vosk/
' Includes Whisper.net in unchanged form; Copyright (c) 2024 Sandro Hanea; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/sandrohanea/whisper.net
' Includes Grpc.core/Grpc.net in unchanged form; Copyright (c) 2023/2025 The gRPC Authors; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/grpc/grpc
' Includes Google Speech V1/V2 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes Google Protobuf in unchanged form; Copyright (c) 2025 Google Inc.; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/protocolbuffers/protobuf
' Includes Google.Api in unchanged form; Copyright (c) 2025 Google LLC; licensed under the BSD-3-Clause license (https://licenses.nuget.org/BSD-3-Clause) at https://github.com/googleapis/gax-dotnet
' Includes Google.Apis in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-api-dotnet-client
' Includes Google.Longrunning in unchanged form; Copyright (c) 2025 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
' Includes MarkdownToRTF in modified form; Copyright (c) 2025 Gustavo Hennig; original licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/GustavoHennig/MarkdownToRtf
' Includes Nito.AsyncEx in unchanged form; Copyright (c) 2021 Stephen Cleary; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/StephenCleary/AsyncEx
' Includes NetOffice libraries in unchanged form; Copyright (c) 2020 Sebastian Lange, Erika LeBlanc; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/netoffice/NetOffice-NuGet
' Includes NAudio.Lame in unchanged form; Copyright (c) 2019 Corey Murtagh; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/Corey-M/NAudio.Lame
' Includes PdfiumViewer in unchanged form; Copyright (c) 2017 Pieter van Ginkel; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/pvginkel/PdfiumViewer
' Includes PDFsharp in unchanged form; Copyright (c) 2025 PDFSharp Team; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://docs.pdfsharp.net/
' Includes System.Interactive.Async in unchanged form; Copyright (c) 2025 by .NET Foundation and Contributors; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/dotnet/reactive
' Includes Microsoft.Playwright (Playwright for .NET) in unchanged form; Copyright (c) 2020 Darío Kondratiuk and other contributors, with modifications copyright (c) Microsoft Corporation where stated in the source files; licensed under the MIT license (https://licenses.nuget.org/MIT) at https://github.com/microsoft/playwright-dotnet
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License, the Microsoft.Web.WebView2 License (for Microsoft.Web.WebView2, see license on https://www.nuget.org/packages/Microsoft.Web.WebView2/ and below) and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, Microsoft.Identity.Client, Microsoft.Identity.Client.Extensions.Msal, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.

' The Word add-in calls the Draw.io online diagramming service/app (https://www.draw.io) via embed.diagrams.net for diagram editing (https://github.com/jgraph/drawio); copyright (c) 2026 by draw.io Ltd and draw.io AG (made available under Apache 2.0 license [https://github.com/jgraph/drawio?tab=Apache-2.0-1-ov-file])

' Licenses of Red Ink and of third-party components and further legal terms/notices are available in the installation folder and via https://redink.ai.
'
' Documentation for developers: See at the end of this file, throughout the code and the manual (https://redink.ai).

Option Explicit On
Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary
Imports System.Globalization
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports System.Net


Partial Public Class ThisAddIn

    Private Const DLLDIAGNOSTICS = False ' Whether to show diagnostics for loading native DLLs

    ' Hardcoded config values

    Public Shared Version As String = "V.210826" & SharedMethods.VersionQualifier
    Public Const AN As String = "Red Ink"
    Public Const AN2 As String = "redink"
    Public Const AN5 As String = "RI" ' for bubble comments 
    Public Const AN6 As String = "Inky" ' for chat

    Private Const ISearch_MinChars = 500         ' minimum characters for a search hit to be relevant
    Private Const ISearch_MaxChars = 4000        ' characters that will be used per search result (rest will be cut off); not applicable to WebView2 retriever
    Private Const ISearch_MaxCrawlErrors = 3     ' maximum number of errors before search is aborted
    Private Const ShortenPercent As Integer = 20
    Private Const SummaryPercent As Integer = 20
    Private Const NetTrigger As String = "(net)"
    Private Const LibTrigger As String = "(lib)"
    Private Const AllTrigger As String = "(all)"
    Private Const BubblesExtractTrigger As String = "(bubbles)"
    Private Const TPMarkupTrigger As String = "(rev)"
    Private Const TPMarkupTriggerL As String = "(rev:"
    Private Const TPMarkupTriggerR As String = ")"
    Private Const TPMarkupTriggerInstruct As String = "(rev[:user])"
    Private Const ExtTrigger As String = "{doc}"
    Private Const ExtDirTrigger As String = "{dir}"
    Private Const ExtUrlTrigger As String = "{url}"
    Private Const ExtTriggerFixed As String = "{[path]}"
    Private Const AddDocTrigger As String = "(adddoc)"
    Private Const MyStyleTrigger As String = "(mystyle)"
    Private Const MultiModelTrigger As String = "(multimodel)"
    Private Const NoFormatTrigger As String = "(noformat)"
    Private Const NoFormatTrigger2 As String = "(nf)"
    Private Const SameAsReplaceTrigger As String = "(sar)"
    Private Const KFTrigger As String = "(keepformat)"
    Private Const KFTrigger2 As String = "(kf)"
    Private Const KPFTrigger As String = "(keepparaformat)"
    Private Const KPFTrigger2 As String = "(kpf)"
    Private Const ObjectTrigger As String = "(file)"
    Private Const ObjectTrigger2 As String = "(clip)"
    Private Const ShowModel As String = "(model)"
    Private Const InPlacePrefix As String = "Replace:"
    Private Const NewdocPrefix As String = "Newdoc:"
    Private Const AddPrefix As String = "Append:"
    Private Const AddPrefix2 As String = "Add:"
    Private Const MarkupPrefix As String = "Markup:"
    Private Const MarkupPrefixDiff As String = "MarkupDiff:"
    Private Const MarkupPrefixDiffW As String = "MarkupDiffW:"
    Private Const MarkupPrefixWord As String = "MarkupWord:"
    Private Const MarkupPrefixRegex As String = "MarkupRegex:"
    Private Const MarkupPrefixAll As String = "Markup[Diff|DiffW|Word|Regex]:"
    Private Const PurePrefix As String = "Pure:"
    Private Const ClipboardPrefix As String = "Clipboard:"
    Private Const ClipboardPrefix2 As String = "Clip:"
    Private Const FilePrefix As String = "File:"
    Private Const FilePrefix2 As String = "Files:"
    Private Const PanePrefix As String = "Pane:"
    Private Const BubblesPrefix As String = "Bubbles:"
    Private Const PushbackPrefix As String = "Reply:"
    Private Const PushbackPrefix2 As String = "Pushback:"
    Private Const SlidesPrefix As String = "Slides:"
    Private Const ChartPrefix As String = "Chart:"
    Private Const ChartPrefixApp As String = "Appchart:"
    Private Const AssemblePrefix As String = "Assemble:"
    Private Const BubbleCutText As String = " (" & ChrW(&H2702) & ")"
    Private Const SearchNextTrigger As String = "Next:"
    Private Const BoWTrigger As String = "(bow)"
    Private Const ChunkTrigger As String = "(iterate)"
    Private Const EmbedTrigger As String = "(embed)"
    Private Const FormPrefix As String = "Form:"
    Private Const RefreshTrigger As String = "(refresh)"
    Private Const ToolSelectionTrigger As String = "(agents)"  ' Trigger in OtherPrompt to re-select tools for tooling-enabled models.
    Private Const ToolTrigger As String = "(ag)"
    Public Const ToolFriendlyName As String = "Agents"  ' How to refer to tools (e.g., sources) towards the user
    Private Const KbTrigger As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTrigger          ' "(kb)"
    Private Const KbTriggerPrefix As String = SharedLibrary.SharedLibrary.KnowledgeTriggerHelper.KbTriggerPrefix ' "(kb:"

    Private Const MaxFilibuster As Integer = 10000 ' Maximum number of words for filibuster mode 
    Private Const ArgueAgainstDefault As Integer = 50 ' Number of words to propose for Argue Against

    Private Const RegexSeparator1 As String = "|||"  ' Set also in SharedLibrary
    Private Const RegexSeparator2 As String = "§§§"  ' Set also in SharedLibrary 
    Private Const RIMenu = AN
    Private Const OldRIMenu = AN & " " & ChrW(&HD83D) & ChrW(&HDC09)
    Private Const MinHelperVersion = 1 ' Minimum version of the helper file that is required

    Public Const IgnoreMarkups As Boolean = False ' Whether to ignore markups in the text when doing a search

    ' MarkupMethod 2 & 5

    ' Set to True to enable optional sentence-level diff switching in the surgical markup engine.
    ' When a sentence is almost entirely rewritten (retention below the threshold, and the sentence
    ' contains no field/footnote/endnote/paragraph-format placeholder and no paragraph/line break),
    ' it is applied as a single tracked delete + insert for cleaner, more readable revisions.
    ' When False, the engine keeps the original word-level behavior fully unchanged.
    Public Shared SurgicalSentenceDiffEnabled As Boolean = True

    ' Maximum fraction of the original sentence's words that may survive for a collapse to occur
    ' (0.2 = collapse only when roughly 80% or more of the words changed).
    Private Const SurgicalSentenceRetentionThreshold As Double = 0.2

    ' Minimum trimmed length (in characters) a sentence must have on both sides before it may be
    ' collapsed. Shorter fragments are already emitted as a single cluster, so collapsing them
    ' would only reduce readability without saving work.
    Private Const SurgicalSentenceMinLength As Integer = 25

    ' Voice Processing

    Private Const VoskSource = "https://alphacephei.com/vosk/models"
    Private Const WhisperSource = "https://huggingface.co/ggerganov/whisper.cpp/tree/main"

    Public Shared WhisperSupportedLanguages As New HashSet(Of String) From {
                            "af", "sq", "am", "ar", "hy", "as", "az", "ba", "eu", "be", "bn", "bs", "br", "bg",
                            "ca", "zh", "hr", "cs", "da", "nl", "en", "et", "fo", "fi", "fr", "gl", "ka", "de",
                            "el", "gu", "ht", "ha", "he", "hi", "hu", "is", "id", "it", "ja", "jv", "kn", "kk",
                            "km", "rw", "ky", "ko", "lv", "lt", "lb", "mk", "mg", "ms", "ml", "mt", "mi", "mr",
                            "mn", "my", "ne", "no", "oc", "ps", "fa", "pl", "pt", "pa", "ro", "ru", "sa", "sr",
                            "sd", "si", "sk", "sl", "so", "es", "su", "sw", "sv", "tl", "tg", "ta", "tt", "te",
                            "th", "tr", "uk", "ur", "uz", "vi", "cy", "yi", "yo", "zu", "auto"
                        }

    Public Shared GoogleTTSsupportedLanguages As String() = {
            "en-US", "en-GB", "de-DE", "fr-FR", "es-ES", "it-IT",
            "af-ZA", "sq-AL", "am-ET", "ar-SA", "eu-ES", "bn-BD",
            "bs-BA", "bg-BG", "yue-HK", "ca-ES", "zh-CN", "zh-TW",
            "hr-HR", "cs-CZ", "da-DK", "nl-NL", "en-AU", "en-IN",
            "en-NG", "et-EE", "fil-PH", "fi-FI", "fr-CA", "gl-ES",
            "el-GR", "gu-IN", "ha-NG", "he-IL", "hi-IN", "hu-HU",
            "is-IS", "id-ID", "ja-JP", "jv-ID", "kn-IN", "km-KH",
            "ko-KR", "la-LA", "lv-LV", "lt-LT", "ms-MY", "ml-IN",
            "mr-IN", "my-MM", "ne-NP", "nb-NO", "pl-PL", "pt-BR",
            "pt-PT", "pa-IN", "ro-RO", "ru-RU", "sr-RS", "si-LK",
            "sk-SK", "es-US", "su-ID", "sw-KE", "sv-SE", "ta-IN",
            "te-IN", "th-TH", "tr-TR", "uk-UA", "ur-PK", "vi-VN", "cy-GB"
        }

    ' Human-readable descriptions for each OpenAI voice.
    Private Shared ReadOnly OpenAIDescriptions As New Dictionary(Of String, String) From {
    {"alloy", "Neutral: balanced and versatile (default, general purpose)"},
    {"echo", "Male: warm, natural, conversational"},
    {"fable", "Male: expressive storyteller (ideal for narration)"},
    {"onyx", "Male: deep, authoritative, strong presence"},
    {"nova", "Female: bright, energetic, modern"},
    {"shimmer", "Female: clear, expressive, polished"}
}


    Private Const TTS_OpenAI_Model = "tts-1-hd"

    Private Shared ReadOnly OpenAIVoices As String() = OpenAIDescriptions.Keys.ToArray()
    Private Shared ReadOnly OpenAILanguages As String() = {
    "de", "en", "es", "fr", "it", "ja", "ko", "pt", "ru", "zh",
    "ar", "bg", "ca", "cs", "da", "el", "et", "fi", "hi", "hu",
    "id", "nl", "no", "pl", "ro", "sv", "th", "tr", "uk", "vi"
}

    Private Const Code_JsonTemplateFormatter As String = "Public Module JsonTemplateFormatter" & vbCrLf & "''' <summary>" & vbCrLf & "''' Hauptfunktion für JSON-String + Template" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(json As String, ByVal template As String) As String" & vbCrLf & "    Dim jObj As JObject" & vbCrLf & "    Try" & vbCrLf & "        jObj = JObject.Parse(json)" & vbCrLf & "    Catch ex As Newtonsoft.Json.JsonReaderException" & vbCrLf & "        Return $""[Fehler beim Parsen des JSON: {ex.Message}]""" & vbCrLf & "    End Try" & vbCrLf & "    NormalizeSources(jObj)" & vbCrLf & "    Return FormatJsonWithTemplate(jObj, template)" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Hauptfunktion für direkten JObject + Template" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function FormatJsonWithTemplate(jObj As JObject, ByVal template As String) As String" & vbCrLf & "    If String.IsNullOrWhiteSpace(template) Then Return """"" & vbCrLf & "    NormalizeSources(jObj)" & vbCrLf & "    ' Normalize CRLF / Platzhalter für Zeilenumbruch" & vbCrLf & "    template = template _" & vbCrLf & "        .Replace(""\\N"", vbCrLf) _" & vbCrLf & "        .Replace(""\\n"", vbCrLf) _" & vbCrLf & "        .Replace(""\\R"", vbCrLf) _" & vbCrLf & "        .Replace(""\\r"", vbCrLf)" & vbCrLf & "    template = Regex.Replace(template, ""<cr>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    Dim hasLoop = Regex.IsMatch(template, ""\\{\\%\\s*for\\s+([^\\s\\%]+)\\s*\\%\\}"", RegexOptions.Singleline)" & vbCrLf & "    Dim hasPh = Regex.IsMatch(template, ""\\{([^}]+)\\}"")" & vbCrLf & "    ' === Einfache Fallbehandlung ===" & vbCrLf & "    If Not hasLoop AndAlso Not hasPh Then" & vbCrLf & "        ' Template enthält keine Platzhalter → als einfacher JSONPath behandeln" & vbCrLf & "        Return FindJsonProperty(jObj, template)" & vbCrLf & "    End If" & vbCrLf & "    ' === Schleifen-Blöcke ===" & vbCrLf & "    Dim loopRegex = New Regex(""\\{\\%\\s*for\\s+([^%\\s]+)\\s*\\%\\}(.*?)\\{\\%\\s*endfor\\s*\\%\\}"", RegexOptions.Singleline Or RegexOptions.IgnoreCase)" & vbCrLf & "    Dim mLoop = loopRegex.Match(template)" & vbCrLf & "    While mLoop.Success" & vbCrLf & "        Dim fullBlock = mLoop.Value" & vbCrLf & "        Dim rawPath = mLoop.Groups(1).Value.Trim()" & vbCrLf & "        Dim innerTpl = mLoop.Groups(2).Value" & vbCrLf & "        Dim path = If(rawPath.StartsWith(""$""), rawPath, ""$."" & rawPath)" & vbCrLf & "        Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "        Dim items = tokens.SelectMany(Function(t)" & vbCrLf & "            If t.Type = JTokenType.Array Then" & vbCrLf & "                Return CType(t, JArray).OfType(Of JObject)()" & vbCrLf & "            ElseIf t.Type = JTokenType.Object Then" & vbCrLf & "                Return {CType(t, JObject)}" & vbCrLf & "            Else" & vbCrLf & "                Return Enumerable.Empty(Of JObject)()" & vbCrLf & "            End If" & vbCrLf & "        End Function)" & vbCrLf & "        Dim rendered = items.Select(Function(o) FormatJsonWithTemplate(o, innerTpl)).ToArray()" & vbCrLf & "        template = template.Replace(fullBlock, If(rendered.Any, String.Join(vbCrLf & vbCrLf, rendered), """"))" & vbCrLf & "        mLoop = loopRegex.Match(template)" & vbCrLf & "    End While" & vbCrLf & "    ' === Platzhalter (non-gierig) ===" & vbCrLf & "    Dim phRegex = New Regex(""\\{(.+?)\\}"", RegexOptions.Singleline)" & vbCrLf & "    Dim result = template" & vbCrLf & "    For Each mPh As Match In phRegex.Matches(template)" & vbCrLf & "        Dim fullPh = mPh.Value" & vbCrLf & "        Dim content = mPh.Groups(1).Value" & vbCrLf & "        ' HTML- oder No-CR-Flag?" & vbCrLf & "        Dim isHtml As Boolean = False" & vbCrLf & "        Dim isNoCr As Boolean = False" & vbCrLf & "        If content.StartsWith(""htmlnocr:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isHtml = True" & vbCrLf & "            isNoCr = True" & vbCrLf & "            content = content.Substring(""htmlnocr:"".Length)" & vbCrLf & "        ElseIf content.StartsWith(""html:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isHtml = True" & vbCrLf & "            content = content.Substring(""html:"".Length)" & vbCrLf & "        ElseIf content.StartsWith(""nocr:"", StringComparison.OrdinalIgnoreCase) Then" & vbCrLf & "            isNoCr = True" & vbCrLf & "            content = content.Substring(""nocr:"".Length)" & vbCrLf & "        End If" & vbCrLf & "        ' Nur am ersten ""|"" trennen" & vbCrLf & "        Dim parts = content.Split(New Char() {""|""c}, 2)" & vbCrLf & "        Dim pathPh = parts(0).Trim()" & vbCrLf & "        Dim remainder = If(parts.Length > 1, parts(1), String.Empty)" & vbCrLf & "        ' Separator-Override (z.B. ""/"") oder Mapping-Definition (enthält ""="")" & vbCrLf & "        Dim sep As String = vbCrLf" & vbCrLf & "        Dim mappings As Dictionary(Of String, String) = Nothing" & vbCrLf & "        If Not String.IsNullOrEmpty(remainder) Then" & vbCrLf & "            If remainder.Contains(""=""c) Then" & vbCrLf & "                mappings = ParseMappings(remainder)" & vbCrLf & "            Else" & vbCrLf & "                sep = remainder.Replace(""\\n"", vbCrLf)" & vbCrLf & "            End If" & vbCrLf & "        End If" & vbCrLf & "        Dim replacement = RenderTokens(jObj, pathPh, sep, isHtml, isNoCr, mappings)" & vbCrLf & "        result = result.Replace(fullPh, replacement)" & vbCrLf & "    Next" & vbCrLf & "    Return result" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Wandelt ausgewählte Tokens in einen String um, wendet Mapping, HTML→Markdown und No-CR an." & vbCrLf & "''' </summary>" & vbCrLf & "Private Function RenderTokens(" & vbCrLf & "    jObj As JObject," & vbCrLf & "    path As String," & vbCrLf & "    sep As String," & vbCrLf & "    isHtml As Boolean," & vbCrLf & "    isNoCr As Boolean," & vbCrLf & "    mappings As Dictionary(Of String, String)" & vbCrLf & ") As String" & vbCrLf & "    Try" & vbCrLf & "        If Not path.StartsWith(""$"") AndAlso Not path.StartsWith(""@"") Then" & vbCrLf & "            path = ""$."" & path" & vbCrLf & "        End If" & vbCrLf & "        Dim tokens = jObj.SelectTokens(path)" & vbCrLf & "        Dim list As New List(Of String)" & vbCrLf & "        For Each t In tokens" & vbCrLf & "            Dim raw = t.ToString()" & vbCrLf & "            ' Mapping anwenden, falls definiert" & vbCrLf & "            If mappings IsNot Nothing AndAlso mappings.ContainsKey(raw) Then raw = mappings(raw)" & vbCrLf & "            ' HTML→Markdown, falls gewünscht" & vbCrLf & "            If isHtml Then raw = HtmlToMarkdownSimple(raw)" & vbCrLf & "            ' No-CR: alle Zeilenumbrüche durch Leerzeichen" & vbCrLf & "            'If isNoCr Then raw = Regex.Replace(raw, ""[\\r\\n]+"", "" "").Trim()" & vbCrLf & "            If isNoCr Then" & vbCrLf & "                ' 1) Turn all line-breaks into single spaces" & vbCrLf & "                raw = Regex.Replace(raw, ""[\\r\\n]+"", "" "")" & vbCrLf & "                ' 2) Collapse any run of whitespace into one space" & vbCrLf & "                raw = Regex.Replace(raw, ""\\s{2,}"", "" "")" & vbCrLf & "                ' 3) Remove common Unicode bullet characters only" & vbCrLf & "                raw = Regex.Replace(raw, ""[\\u2022\\u2023\\u25E6]"", String.Empty)" & vbCrLf & "                ' 4) Trim leading/trailing spaces" & vbCrLf & "                raw = raw.Trim()" & vbCrLf & "            End If" & vbCrLf & "            list.Add(raw)" & vbCrLf & "        Next" & vbCrLf & "        Return If(list.Count = 0, """", String.Join(sep, list))" & vbCrLf & "    Catch ex As System.Exception" & vbCrLf & "        Return """"" & vbCrLf & "    End Try" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Parst Mapping-Definitionen der Form ""key1=Text1;key2=Text2;…""" & vbCrLf & "''' </summary>" & vbCrLf & "Private Function ParseMappings(defs As String) As Dictionary(Of String, String)" & vbCrLf & "    Dim dict As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)" & vbCrLf & "    For Each pair In defs.Split("";""c)" & vbCrLf & "        Dim kv = pair.Split(New Char() {""=""c}, 2)" & vbCrLf & "        If kv.Length = 2 Then dict(kv(0).Trim()) = kv(1).Trim()" & vbCrLf & "    Next" & vbCrLf & "    Return dict" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "''' <summary>" & vbCrLf & "''' Einfacher HTML→Markdown-Konverter (inkl. SPAN → *italic*)" & vbCrLf & "''' </summary>" & vbCrLf & "Public Function HtmlToMarkdownSimple(html As String) As String" & vbCrLf & "    Dim s = WebUtility.HtmlDecode(html)" & vbCrLf & "    ' Absätze → zwei Zeilenumbrüche            " & vbCrLf & "    s = Regex.Replace(s, ""</?p\\s*/?>"", vbCrLf & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Zeilenumbruch-Tags" & vbCrLf & "    s = Regex.Replace(s, ""<br\\s*/?>"", vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Fett/strong → **text**" & vbCrLf & "    s = Regex.Replace(s, ""<strong>(.*?)</strong>"", ""**$1**"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Kursiv/em → *text*" & vbCrLf & "    s = Regex.Replace(s, ""<em>(.*?)</em>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' SPAN-Tags → *text*" & vbCrLf & "    s = Regex.Replace(s, ""<span\\b[^>]*>(.*?)</span>"", ""*$1*"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Listenpunkte <li> → ""- text""" & vbCrLf & "    s = Regex.Replace(s, ""<li>(.*?)</li>"", ""- $1"" & vbCrLf, RegexOptions.IgnoreCase)" & vbCrLf & "    ' Fußnoten-Tags <fn>…</fn> → <sup>…</sup>" & vbCrLf & "    s = Regex.Replace(s, ""<fn>(.*?)</fn>"", ""<sup>$1</sup>"", RegexOptions.IgnoreCase)" & vbCrLf & "    ' Alle übrigen Tags entfernen" & vbCrLf & "    s = Regex.Replace(s, ""<(?!/?sup\\b)[^>]+>"", String.Empty, RegexOptions.IgnoreCase)" & vbCrLf & "    's = Regex.Replace(s, ""<[^>]+>"", String.Empty)" & vbCrLf & "    ' Mehrfache Zeilenumbrüche aufräumen" & vbCrLf & "    s = Regex.Replace(s, ""("" & vbCrLf & ""){3,}"", vbCrLf & vbCrLf)" & vbCrLf & "    Return s.Trim()" & vbCrLf & "End Function" & vbCrLf & "" & vbCrLf & "Private Sub NormalizeSources(jObj As JObject)" & vbCrLf & "    Dim srcToken = jObj.SelectToken(""sources"")" & vbCrLf & "    If srcToken IsNot Nothing AndAlso srcToken.Type = JTokenType.Array Then" & vbCrLf & "        Dim newArray As New JArray()" & vbCrLf & "        For Each item In CType(srcToken, JArray)" & vbCrLf & "            If item.Type = JTokenType.Array AndAlso item.Count >= 3 Then" & vbCrLf & "                Dim objStr = item(2).ToString()" & vbCrLf & "                Try" & vbCrLf & "                    Dim o = JObject.Parse(objStr)" & vbCrLf & "                    newArray.Add(o)" & vbCrLf & "                Catch ex As System.Exception" & vbCrLf & "                    ' Ungültiges JSON überspringen" & vbCrLf & "                End Try" & vbCrLf & "            ElseIf item.Type = JTokenType.Object Then" & vbCrLf & "                newArray.Add(item)" & vbCrLf & "            End If" & vbCrLf & "        Next" & vbCrLf & "        jObj(""sources"") = newArray" & vbCrLf & "    End If" & vbCrLf & "End Sub" & vbCrLf & "" & vbCrLf & "End Module"

    Private Const SP_GenerateResponseKey As String = "I have code that will generate from a JSON string an Markdown output using a Template, which the code will parse together with the JSON file. I want you to create me a working template taking into account (i) the code, (ii) the structure of the JSON file and (iii) my instructions. If the JSON has arrays, make sure you correctly handle them. To produce your output, first provide the barebones template one one single line (do not use placeholders, provide the text how template should look like; for linebreaks, use only <cr>), then provide a brief explanation without any formatting. I will provide you in the following first the code, and then you will get the (sample) JSON file and my instructions. Follow them carefully."

    Private Const NER_Model = "anon\model.onnx"
    Private Const NER_Token = "anon\bpe.model"
    Private Const NER_Label = "anon\label_map.txt"
    Private Const Embed_Model = "embed\model.onnx"
    Private Const Embed_Vocab = "embed\vocab.txt"
    Private Default_Embed_Min_Score As Double = 0.2
    Private Default_Embed_Top_K As Integer = 5
    Private Default_Embed_Chunks As Integer = 2
    Private Default_Embed_Overlap As Integer = 1
    Private Default_Embed_Chunks_bow As Integer = 1
    Private Default_Embed_Overlap_bow As Integer = 0

    Public Shared DragDropFormLabel As String = ""
    Public Shared DragDropFormFilter As String = ""

    Public Shared TTSDefaultFile As String = $"{AN2}-output.mp3"
    Public Const TTSLargeText As Integer = 2500
    Public Shared hostTags As String() = {"H:", "Host:", "A:", "1:"}
    Public Shared guestTags As String() = {"G:", "Guest:", "Gast:", "B:", "2:"}
    Public Shared GoogleIdentifier As String = "googleapis.com"
    Public Shared OpenAIIdentifier As String = "openai.com"

    Public Shared TTS_googleAvailable As Boolean = False
    Private Shared TTS_googleSecondary As Boolean = False
    Public Shared TTS_openAIAvailable As Boolean = False
    Private Shared TTS_openAISecondary As Boolean = False
    Private Shared TTS_GoogleEndpoint As String = ""
    Private Shared TTS_OpenAIEndpoint As String = ""

    ' Tooling

    Public Const ToolingLog_AutoCloseDefaultSeconds As Integer = 20

    Public Const InternalToolSuffix As String = ""  ' Suffix displayed for the internal web tool in selection dialogs.

    Public Const InternalWebToolName As String = "web_content_retriever"
    Public Const InternalWebToolDescription As String =
        "Retrieves readable text from one or more web pages. By default it returns text only. " &
        "Optionally, the caller can request structured link extraction for any linked content."

    Public Const InternalWebToolDefinition As String =
        "{""name"":""web_content_retriever"",""description"":""Fetches and returns readable text from one or more web URLs. " &
        "By default this tool returns text only. " &
        "Optional behavior: set include_links=true to also extract structured hyperlinks; set expand_interactive_sections=true " &
        "to attempt opening accordions, tabs, details elements, and similar collapsed UI before extraction; set link_extensions " &
        "to filter extracted links to specific extensions such as pdf, docx, xlsx, zip, html, or any other extension. " &
        "IMPORTANT: Cannot access SharePoint, OneDrive, Teams, or other authenticated cloud storage URLs " &
        "(sharepoint.com, onedrive.com, 1drv.ms, teams.microsoft.com, :f:/). " &
        "Do NOT call this tool for such URLs — ask the user to download and attach the file(s) instead."",""parameters"":{""type"":""object"",""properties"":{" &
        """urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""One or more absolute URLs to fetch (preferred).""}," &
        """url"":{""type"":""string"",""description"":""Single absolute URL to fetch (alternative to urls).""}," &
        """include_links"":{""type"":""boolean"",""description"":""Optional. Default false. When true, include a structured <LINKS_n> JSON block for each URL.""}," &
        """link_extensions"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Optional. Restrict extracted links to these extensions, for example ['pdf','docx']. Ignored unless include_links=true.""}," &
        """expand_interactive_sections"":{""type"":""boolean"",""description"":""Optional. Default false for tooling calls. When true, attempts to expand accordions, tabs, details sections, and similar collapsed UI before extraction.""}}}}" ' Runtime validates url/urls.

    Public Const InternalWebToolInstructionsPrompt As String =
        "web_content_retriever: Fetches readable text from web pages. " &
        "Call this tool when you need the actual page content behind a link. " &
        "Provide either urls (array of strings) or url (single string). " &
        "By default, this tool returns readable text only. " &
        "If you also need structured links to downloadable or linked content, explicitly set include_links=true. " &
        "If linked content may be hidden behind accordions, tabs, or collapsed sections, explicitly set expand_interactive_sections=true. " &
        "Optionally set link_extensions to restrict extracted links to specific extensions such as ['pdf'], ['docx'], or any other extension. " &
        "Return value is plain text content for each URL, optionally followed by a structured <LINKS_n> JSON block for each URL. " &
        "IMPORTANT BOUNDARY: This tool retrieves readable text and link metadata only. It does NOT download or preserve original binary file bytes. " &
        "If you call this tool on a direct PDF, DOCX, XLSX, PPTX, ZIP, image, audio, or other binary URL, the result is text extraction or page content — not a real downloadable file object. " &
        "Therefore, NEVER use this tool as a file-downloader and NEVER save its returned text as if it were the original PDF or other binary file. " &
        "If the user wants the actual remote file saved into the workspace, use a dedicated binary download tool if available; otherwise explain that the current toolset can analyze the content but cannot save the original remote binary file. " &
        "IMPORTANT FALLBACK RULE: If include_links=true returns zero matching links, and the page may be dynamically computing or revealing links client-side, then use js_run as a fallback with allow_network=true and navigate_url set to that page. " &
        "In js_run, the code is already the body of an async function, so it must return the final value explicitly at top level. " &
        "SHAREPOINT/ONEDRIVE LIMITATION: This tool CANNOT access SharePoint, OneDrive, Microsoft Teams, or any other " &
        "authenticated cloud storage URLs. URLs containing 'sharepoint.com', 'onedrive.com', '1drv.ms', " &
        "'teams.microsoft.com', or ':f:/' point to resources that require authentication and will NOT return " &
        "useful content. UNC paths (e.g. \\server\share\file.doc) that resolve to SharePoint will also fail. " &
        "If the user asks you to retrieve content from such a link, do NOT call this tool. Instead, explain " &
        "that you cannot remotely log into authenticated cloud storage and ask the user to download the file(s) " &
        "and provide them as direct attachments."

    Public Const InternalDownloadWebFilesToolName As String = "download_web_files"

    Public Const InternalDownloadWebFilesToolDefinition As String =
        "{""name"":""download_web_files"",""description"":""Downloads one or more remote files and saves the original binary bytes to disk. " &
        "Use this for real file downloads, not for text extraction. Only absolute HTTP/HTTPS URLs are allowed. " &
        "Authenticated SharePoint, OneDrive, Teams, and similar cloud storage URLs are not supported.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""Array of absolute HTTP/HTTPS URLs to download.""}," &
        """target_directory"":{""type"":""string"",""description"":""Optional relative target folder. If a writable workspace is active, this is resolved under the workspace root; otherwise under the safe download root.""}," &
        """overwrite"":{""type"":""boolean"",""description"":""Optional. Default false. If false, existing files are not overwritten; unique names are created instead.""}}" &
        ",""required"":[""urls""]}}"

    Public Const InternalDownloadWebFilesToolInstructionsPrompt As String =
        "download_web_files: Downloads one or more remote files and saves the original binary bytes to disk. " &
        "Use this only when the user wants the actual file saved locally. " &
        "Provide urls (required array). Optionally provide target_directory (relative subfolder under the current workspace when available; otherwise under the safe download root) and overwrite (boolean, default false). " &
        "Use this tool instead of web_content_retriever when the user wants the actual PDF or other binary file, not merely extracted text."

    ' Internet Search Tooling (available only when INI_ISearch is enabled and INI_ISearch_URL is configured)

    Public Const InternalSearchToolName As String = "internet_search"

    Public Const InternalSearchToolDefinition As String =
        "{""name"":""internet_search""," &
        """description"":""Searches the internet via the configured search engine, retrieves the top result pages, and returns their readable text content. Use this when you need up-to-date or factual information you are not confident about. PRIVACY: The query is sent to an external search engine. Never include personal data, confidential information, private names, case details, contract terms, internal identifiers, or any non-public information in the query. Only public figures, public institutions, published legislation, and other clearly public information may appear. If a useful query cannot be formed without non-public data, do not call this tool.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""The search query. MUST NOT contain personal data, confidential details, or any non-public information. Use only generic, anonymized, or publicly known terms.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of search result pages to retrieve (default: 4, server-capped).""}," &
        """max_depth"":{""type"":""integer"",""description"":""Maximum crawl depth per result page. 0 = top-level only (default: 0, server-capped).""}},""required"":[""query""]}}"

    Public Const InternalSearchToolInstructionsPrompt As String =
        "internet_search: Searches the internet and returns readable text from the top result pages. " &
        "Call this tool when you need current or factual information you are not confident about. " &
        "Provide query (required string). Optionally provide max_results (integer, default 4) and max_depth (integer, default 0). " &
        "Return value includes the search query used, the URLs visited, and the page content for each qualifying result. " &
        "IMPORTANT PRIVACY CONSTRAINT: The search query is sent to an external search engine. " &
        "You MUST NOT include any personal data, confidential information, private names, " &
        "case details, contract terms, internal identifiers, email addresses, phone numbers, " &
        "account numbers, or any other non-public information in the query. " &
        "Only well-known public figures, public institutions, published legislation, " &
        "publicly available case law references, and other clearly public information may appear in queries. " &
        "If you cannot formulate a useful query without disclosing non-public information, " &
        "do NOT call this tool — instead respond based on your existing knowledge and state your uncertainty."


    ' Knowledge Store Tooling (available only when KnowledgeStorePath or KnowledgeStorePathLocal is configured)

    Public Const InternalKnowledgeToolName As String = "knowledge_search"

    Public Const InternalKnowledgeToolDefinition As String =
        "{""name"":""knowledge_search""," &
        """description"":""Searches the user's local knowledge store (a curated collection of documents such as contracts, policies, legal briefs, " &
        "manuals, and reference material) and returns the most relevant document content. Use this tool when the user's question " &
        "relates to their own documents, internal policies, past work, or reference material that would not be found on the public internet. " &
        "Do NOT use this tool for general knowledge questions or publicly available information — use your training data or internet_search instead.""," &
        """parameters"":{""type"":""object"",""properties"":{" &
        """query"":{""type"":""string"",""description"":""A natural language search query describing what information is needed from the knowledge store. " &
        "Supports optional prefixes: 'tag:tagname' to filter by tag, 'store:storename' to restrict to a specific store, " &
        "or both 'tag:tagname store:storename'. Without prefixes, all stores are searched by keyword relevance.""}," &
        """max_results"":{""type"":""integer"",""description"":""Maximum number of documents to retrieve (default: 5, max: 10).""}},""required"":[""query""]}}"

    Public Const InternalKnowledgeToolInstructionsPrompt As String =
        "knowledge_search: Searches the user's local knowledge store — a curated library of the user's own documents " &
        "(contracts, policies, briefs, manuals, templates, reference material, etc.). " &
        "Call this tool when the user's question relates to their own documents, internal policies, past work, or reference material. " &
        "Provide query (required string) describing the information needed. Optionally provide max_results (integer, default 5). " &
        "The query supports optional prefixes: 'tag:tagname' filters by document tag, 'store:storename' restricts to a specific knowledge store. " &
        "Return value is the text content of the most relevant documents, each tagged with the document name and store. " &
        "IMPORTANT: Do NOT use this tool for general knowledge or publicly available information — only for the user's own document library. " &
        "When citing information from the results, mention the document name so the user can locate the source."

    Public Shared SelectedToolNames As New List(Of String)()   ' Persisted list of selected tool names for tooling sessions.


    Private Const Add_OcrMarkdownInstruction As String =
            "Provide the output as Markdown. " &
            "Use Markdown formatting where it helps preserve the original structure, including headings, lists, tables and other tabular content. " &
            "Return only the Markdown content and no surrounding code fences."



    ' Declare variables publicly so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public SourceLanguage As String
    Public ShortenLength As Double
    Public FilibusterLength As Integer
    Public SummaryLength As Integer
    Public OtherPrompt As String = ""
    Public OtherPromptUnfilled As String = ""
    Public Dictionary As String = ""
    Public OutputLanguage As String = ""

    Public ReadOnly Property MaxToolIterations As Integer
        Get
            Return INI_ToolingMaximumIterations
        End Get
    End Property

    Public InsertDocs As String = ""
    Public MyStyleInsert As String = ""
    Public FormatInstruction As String = ""
    Public SearchTerms As String
    Public SearchContext As String

    Public ReadOnly Property CurrentDate As String
        Get
            Return "Current date (ISO 8601 yyyy-MM-dd): " &
                DateTime.Now.ToString("yyyy-MM-dd", Globalization.CultureInfo.InvariantCulture)
        End Get
    End Property

    Public ReadOnly Property LocalTimeContext As String
        Get
            Dim localNow As System.DateTimeOffset = System.DateTimeOffset.Now
            Dim localTz As System.TimeZoneInfo = System.TimeZoneInfo.Local
            Dim localZoneName As String = If(
                localTz.IsDaylightSavingTime(System.DateTime.Now),
                localTz.DaylightName,
                localTz.StandardName)

            Return "Current local runtime time (ISO 8601): " &
                localNow.ToString("yyyy-MM-dd HH:mm:ss zzz", Globalization.CultureInfo.InvariantCulture) &
                $". Local timezone: {localZoneName} ({localTz.Id}). Unless the user explicitly asks for UTC or another timezone, use this local timezone for user-facing dates and times."
        End Get
    End Property

    Public SysPrompt As String
    Public OldParty, NewParty As String
    Public SelectedText As String
    Public LibraryText As String
    Public LibResult As String
    Public SearchResult As String
    Public doc As String
    Public HostName As String
    Public GuestName As String
    Public Language As String
    Public Duration As String
    Public TargetAudience As String
    Public DialogueContext As String
    Public ExtraInstructions As String
    Public DiscussKnowledgeCache As String = ""


    <DllImport("user32.dll", SetLastError:=True)>
    Private Shared Function FindWindow(
                                ByVal lpClassName As String,
                                ByVal lpWindowName As String
                            ) As IntPtr
    End Function

    Private Function GetWordMainWindowHandle() As IntPtr
        ' Word’s top-level windows all have the class name "OpusApp" (Office 2013+)
        Dim hwnd = FindWindow("OpusApp", Nothing)
        Return hwnd
    End Function

    Private mainThreadControl As New System.Windows.Forms.Control()
    Public StartupInitialized As Boolean = False
    Private WithEvents wordApp As Word.Application

    ' UI threading context and scheduler (captured at Startup)
    Private Shared _uiContext As SynchronizationContext
    Private Shared _uiScheduler As TaskScheduler
    ''' <summary>
    ''' Captured WindowsFormsSynchronizationContext for the Word UI/STA thread.
    ''' Used to marshal LogWindow creation and updates back to the UI thread,
    ''' regardless of whether the caller awaited with ConfigureAwait(False).
    ''' </summary>
    Public Shared UiSyncContext As System.Threading.SynchronizationContext
    Public Shared UiThreadId As Integer


    Private Shared Sub EnsureModernTls()
        Try
            AppContext.SetSwitch("Switch.System.Net.DontEnableSchUseStrongCrypto", False)
        Catch
        End Try

        Try
            AppContext.SetSwitch("Switch.System.Net.DontEnableSystemDefaultTlsVersions", False)
        Catch
        End Try

        Try
            ServicePointManager.Expect100Continue = False
        Catch
        End Try

        Try
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
        Catch
            Try
                ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)
            Catch
            End Try
        End Try
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Crash diagnostics: enabled only when the user's My.Settings.CrashLog is True.
        ' This flag is reconciled from INI_Crashlog after config load, so it takes effect
        ' on the next host launch. When False, no handlers are installed (zero overhead).
        Try
            RiCrashLogger.Initialize(
                "RedInk Word Add-in",
                Me.GetType().Assembly,
                My.Settings.CrashLog,
                True,
                RDV)
        Catch
        End Try

        ' Necessary for Update Handler to work correctly

        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' Ensure a WinForms sync context exists on the UI thread, then capture it.
        If System.Threading.SynchronizationContext.Current Is Nothing Then
            System.Threading.SynchronizationContext.SetSynchronizationContext(
        New System.Windows.Forms.WindowsFormsSynchronizationContext())
        End If
        UiSyncContext = System.Threading.SynchronizationContext.Current
        UiThreadId = System.Threading.Thread.CurrentThread.ManagedThreadId

        ' 2) Capture synchronization context & scheduler after handle exists
        _uiContext = SynchronizationContext.Current
        If _uiContext Is Nothing Then
            _uiContext = New WindowsFormsSynchronizationContext()
            SynchronizationContext.SetSynchronizationContext(_uiContext)
        End If
        _uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()

        ' 3) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 4) Capture the host window’s HWND.
        Dim hwnd As IntPtr = GetWordMainWindowHandle()
        UpdateHandler.HostHandle = hwnd

        ' Other tasks that need to be done at startup

        EnsureModernTls()

        SharedMethods.Initialize(Me.CustomTaskPanes)

        SharedLibrary.Agents.WordHostPolicy.Host = Me

        If System.Threading.SynchronizationContext.Current Is Nothing Then
            System.Threading.SynchronizationContext.SetSynchronizationContext(
        New System.Windows.Forms.WindowsFormsSynchronizationContext())
        End If

        wordApp = Application
        Try
            If wordApp IsNot Nothing Then
                AddHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
                AddHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
                AddHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
                AddHandler wordApp.ProtectedViewWindowOpen, AddressOf WordApp_ProtectedViewWindowOpen
                AddHandler wordApp.ProtectedViewWindowBeforeEdit, AddressOf WordApp_ProtectedViewWindowBeforeEdit
                AddHandler wordApp.ProtectedViewWindowActivate, AddressOf WordApp_ProtectedViewWindowActivate
                AddHandler wordApp.DocumentChange, AddressOf WordApp_DocumentChange
            Else
                mainThreadControl.BeginInvoke(CType(AddressOf DelayedStartupTasks, MethodInvoker))
                StartupInitialized = True
            End If
            If wordApp.Documents.Count > 0 Then
                'Run everything on the Office UI thread
                mainThreadControl.BeginInvoke(
                                            Sub()
                                                'Detach the one-shot startup hooks
                                                RemoveStartupHandlers()          'sets StartupInitialized = True
                                                DelayedStartupTasks()
                                            End Sub)
            End If
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Sub RemoveStartupHandlers()
        StartupInitialized = True
        Try
            RemoveHandler wordApp.WindowActivate, AddressOf WordApp_WindowActivate
            RemoveHandler wordApp.DocumentOpen, AddressOf WordApp_DocumentOpen
            RemoveHandler wordApp.NewDocument, AddressOf WordApp_NewDocument
            RemoveHandler wordApp.ProtectedViewWindowOpen, AddressOf WordApp_ProtectedViewWindowOpen
            RemoveHandler wordApp.ProtectedViewWindowBeforeEdit, AddressOf WordApp_ProtectedViewWindowBeforeEdit
            RemoveHandler wordApp.ProtectedViewWindowActivate, AddressOf WordApp_ProtectedViewWindowActivate
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Shared Function EnsureUIThread() As Task
        ' If UI context not set or already on it, nothing to do.
        If _uiContext Is Nothing OrElse SynchronizationContext.Current Is _uiContext Then
            Return Task.CompletedTask
        End If
        Dim tcs As New TaskCompletionSource(Of Object)(TaskCreationOptions.RunContinuationsAsynchronously)
        _uiContext.Post(
            Sub(state As Object)
                tcs.TrySetResult(Nothing)
            End Sub,
            Nothing)
        Return tcs.Task
    End Function


    Private Sub WordApp_WindowActivate(ByVal Doc As Word.Document, ByVal Wn As Word.Window)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_DocumentOpen(doc As Word.Document)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_NewDocument(doc As Word.Document)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub


    ' Fires when a file opens in Protected View.
    Private Sub WordApp_ProtectedViewWindowOpen(
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires just before the user clicks “Edit” in Protected View.
    Private Sub WordApp_ProtectedViewWindowBeforeEdit(
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow,
            ByRef Cancel As Boolean)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    ' Fires when the Protected View window is activated.
    Private Sub WordApp_ProtectedViewWindowActivate(
            pvWin As Microsoft.Office.Interop.Word.ProtectedViewWindow)
        RemoveStartupHandlers()
        DelayedStartupTasks()
    End Sub

    Private Sub WordApp_DocumentChange()
        If Not StartupInitialized Then
            RemoveStartupHandlers()
            DelayedStartupTasks()
        End If
    End Sub

    Private Sub DelayedStartupTasks()
        Try
            InitializeAddInFeatures()
            StartupHttpListener(INI_WebServerBlock)
            ' Initialize Knowledge Store background indexing service
            InitializeKnowledgeStoreService()

            Try
                If System.Threading.SynchronizationContext.Current Is Nothing Then
                    System.Threading.SynchronizationContext.SetSynchronizationContext(
                        New System.Windows.Forms.WindowsFormsSynchronizationContext())
                End If
                ' Touch a Control on the UI thread so the WindowsForms sync-context is fully active.
                Using anchor As New System.Windows.Forms.Control()
                    Dim h = anchor.Handle
                End Using
                SharedLibrary.SharedLibrary.SharedMethods.CleanupOrphanedWebView2Profiles()
                SharedLibrary.Agents.WebView2JsSandbox.Initialize(
                    System.Threading.SynchronizationContext.Current)
            Catch
                ' js_run will report "sandbox_uninitialized" if this failed.
            End Try

            ' Warm the Python Agent version cache off the UI thread so the first tooling run
            ' does not pay the version-probe cost. The lazy path in ResolveAndValidateAvailability
            ' re-probes if this warm-up is skipped or the cache is cold.
            PrimePythonAgentVersionCache()

        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    ''' <summary>
    ''' Primes context-independent model/tool INI parsing and the skill/agent index off the
    ''' Word UI thread. The on-demand paths remain authoritative and retry on failure.
    ''' No Office COM or WinForms objects are touched by the background task.
    ''' </summary>
    Private Sub QueueModelAndAgentResourceWarmup()
        Try
            Dim alternateModelPath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AlternateModelPath, ""))
            Dim specialServicePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_SpecialServicePath, ""))
            Dim centralResourcePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AgentResourcesPath, ""))
            Dim localResourcePath As String = SharedMethods.ExpandEnvironmentVariables(If(INI_AgentResourcesPathLocal, ""))

            SharedLibrary.Agents.AgentResources.SetPaths(centralResourcePath, localResourcePath)

            System.Threading.Tasks.Task.Run(
                Sub()
                    Dim stopwatch As System.Diagnostics.Stopwatch = System.Diagnostics.Stopwatch.StartNew()
                    Try
                        SharedMethods.WarmAlternativeModelsCache(alternateModelPath)
                        SharedMethods.WarmAlternativeModelsCache(specialServicePath)
                        SharedLibrary.Agents.AgentResources.EnsureFresh()
                    Catch ex As System.Exception
                        System.Diagnostics.Debug.WriteLine("[PERF] Word startup warm-up failed: " & ex.Message)
                    Finally
                        stopwatch.Stop()
                        System.Diagnostics.Debug.WriteLine(
                            "[PERF] Word model/tool/resource warm-up: " &
                            stopwatch.ElapsedMilliseconds.ToString(System.Globalization.CultureInfo.InvariantCulture) &
                            " ms")
                    End Try
                End Sub)
        Catch ex As System.Exception
            ' Non-fatal. Each on-demand path performs the same work if the cache is cold.
        End Try
    End Sub

    ''' <summary>
    ''' Fire-and-forget warm-up of the Python Agent version cache. Runs off the UI thread and never
    ''' throws; on failure the on-demand tool-registration path re-probes.
    ''' </summary>
    Private Sub PrimePythonAgentVersionCache()
        System.Threading.Tasks.Task.Run(
            Sub()
                Try
                    Dim configuration = SharedLibrary.Agents.PythonExecuteToolConfig.Parse(
                        SharedMethods.ExpandEnvironmentVariables(INI_PythonAgentPath))
                    If configuration IsNot Nothing Then
                        SharedLibrary.Agents.PythonExecuteToolCore.PrimeAvailabilityCache(configuration)
                    End If
                Catch ex As System.Exception
                    ' Non-fatal: the tool-registration path will re-probe on demand.
                End Try
            End Sub)
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        Try
            RiCrashLogger.Shutdown("ThisAddIn_Shutdown was called.")
        Catch
        End Try
        ShutdownTalkToMe()
        ShutdownKnowledgeStoreService()
        Try
            SharedLibrary.Agents.BrowserTools.Shutdown()
        Catch ex As System.Exception
            System.Diagnostics.Debug.WriteLine("BrowserTools.Shutdown failed: " & ex.Message)
        End Try
        ShutdownHttpListener()
        RemoveOldContextMenu()
    End Sub

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)
        QueueModelAndAgentResourceWarmup()

        ' Reconcile the persisted CrashLog switch with the INI parameter. Any change
        ' takes effect on the next host launch (the INI is read after ThisAddIn_Startup).
        Try
            If My.Settings.CrashLog <> INI_Crashlog Then
                My.Settings.CrashLog = INI_Crashlog
                My.Settings.Save()
            End If
        Catch
        End Try

        If DLLDIAGNOSTICS Then WriteDllLoadDiagnosticsIfEnabled()

        ' Restore the previously selected primary model (if multi-model is configured)
        If _context.INIloaded Then
            Try
                Dim saved = PrimaryModelManager.LoadSavedModelNumber()
                If PrimaryModelManager.GetAvailableModels().Count > 0 Then
                    ' Try saved selection first; fall back to model 1 if it no longer exists
                    If Not PrimaryModelManager.SelectModel(_context, saved) Then
                        PrimaryModelManager.SelectModel(_context, PrimaryModelManager.GetAvailableModels()(0))
                    End If
                End If
            Catch
                ' non-critical
            End Try
        End If

        AddContextMenu()

        If _context IsNot Nothing AndAlso _context.INIloaded Then
            SharedMethods.RegisterLicenseCounterUsageAndReport(_context, "WD")
        End If

        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath, _context)

        ' Initialize model menu buttons on the ribbon
        Try
            If Globals.Ribbons.Ribbon1 IsNot Nothing Then
                Globals.Ribbons.Ribbon1.UpdateModelsMenu()
                Globals.Ribbons.Ribbon1.ApplyRibbonVisibilityConfiguration()
            End If
        Catch
            ' non-critical
        End Try

    End Sub


    ' Bridge to SharedLibrary
    Public Sub InitializeConfig(FirstTime As Boolean, Reload As Boolean)
        _context.InitialConfigFailed = False
        _context.RDV = "Word (" & Version & ")"
        SharedMethods.InitializeConfig(_context, FirstTime, Reload)
    End Sub
    Private Function INIValuesMissing() As Boolean
        Return SharedMethods.INIValuesMissing(_context)
    End Function
    Public Shared Async Function PostCorrection(inputText As String, Optional ByVal UseSecondAPI As Boolean = False) As Task(Of String)
        Return Await SharedMethods.PostCorrection(_context, inputText, UseSecondAPI)
    End Function
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "", Optional ByVal ToolExecution As Boolean = False, Optional cancellationToken As Threading.CancellationToken = Nothing, Optional EnsureUI As Boolean = True) As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt, FileObject, cancellationToken, ToolExecution:=ToolExecution)
        If EnsureUI Then
            Await EnsureUIThread().ConfigureAwait(False)
        End If
        Return Response
    End Function
    Private Sub ShowSettingsWindow(Settings As Dictionary(Of String, String), SettingsTips As Dictionary(Of String, String))
        SharedMethods.ShowSettingsWindow(Settings, SettingsTips, _context)
    End Sub
    Private Function ShowPromptSelector(filePath As String, filePathlocal As String, enableMarkup As Boolean, enableBubbles As Boolean) As (String, Boolean, Boolean, Boolean)
        Return SharedMethods.ShowPromptSelector(filePath, filePathlocal, enableMarkup, enableBubbles, _context)
    End Function


    Public Enum CustomWdKey
        wdKeyUp = 38
        wdKeyDown = 40
        wdKeyLeft = 37
        wdKeyRight = 39
        wdKeySpace = 32
    End Enum

    Private automationObject As BridgeSubs

    Protected Overrides Function RequestComAddInAutomationService() As Object
        If automationObject Is Nothing Then
            automationObject = New BridgeSubs()
        End If
        Return automationObject
    End Function

    Private Sub WriteDllLoadDiagnosticsIfEnabled()
        Try
            If _context Is Nothing Then Return
            If Not _context.INIloaded Then
                ShowCustomMessageBox("Cannot generate RI DLL load diagnostics because the configuration could not be loaded.")
                Return
            End If
            If Not _context.INI_APIDebug Then Return

            Dim diagnosticsDirectory As String =
                System.IO.Path.GetDirectoryName(RiCrashLogger.LogFilePath)

            If String.IsNullOrWhiteSpace(diagnosticsDirectory) Then
                diagnosticsDirectory =
                    System.IO.Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "redink")
            End If

            If String.IsNullOrWhiteSpace(diagnosticsDirectory) Then
                ShowCustomMessageBox("Cannot generate RI DLL load diagnostics because the diagnostics path is invalid.")
                Return
            End If

            Try
                System.IO.Directory.CreateDirectory(diagnosticsDirectory)
            Catch ex As Exception
                ShowCustomMessageBox("Cannot generate RI DLL load diagnostics because the diagnostics directory could not be created.")
                Return
            End Try

            Dim outputPath As String = System.IO.Path.Combine(diagnosticsDirectory, "RI_DLL_Loaded.txt")
            Dim report As New System.Text.StringBuilder()

            report.AppendLine("RI DLL Loaded Diagnostic Report")
            report.AppendLine(New String("="c, 80))
            report.AppendLine("Created: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff", CultureInfo.InvariantCulture))
            report.AppendLine("App: " & AN & " for Word")
            report.AppendLine("Red Ink Version: " & Version)
            report.AppendLine("RDV: " & If(_context.RDV, ""))
            report.AppendLine("Machine: " & Environment.MachineName)
            report.AppendLine("User: " & Environment.UserName)
            report.AppendLine("OS Version: " & Environment.OSVersion.ToString())
            report.AppendLine(".NET Version: " & Environment.Version.ToString())
            report.AppendLine("64-bit OS: " & Environment.Is64BitOperatingSystem.ToString())
            report.AppendLine("64-bit Process: " & Environment.Is64BitProcess.ToString())
            report.AppendLine("Current Directory: " & SafeValue(Environment.CurrentDirectory))
            report.AppendLine("Base Directory: " & SafeValue(AppDomain.CurrentDomain.BaseDirectory))
            report.AppendLine("AppDomain Config File: " & SafeValue(AppDomain.CurrentDomain.SetupInformation.ConfigurationFile))
            report.AppendLine("API Debug Enabled: " & _context.INI_APIDebug.ToString())
            report.AppendLine()

            Try
                Dim currentProcess = System.Diagnostics.Process.GetCurrentProcess()
                report.AppendLine("Process Name: " & SafeValue(currentProcess.ProcessName))
                report.AppendLine("Process Id: " & currentProcess.Id.ToString(CultureInfo.InvariantCulture))
                Try
                    report.AppendLine("Process Path: " & SafeValue(currentProcess.MainModule.FileName))
                Catch ex As Exception
                    report.AppendLine("Process Path: <error: " & ex.GetType().FullName & ": " & ex.Message & ">")
                End Try
            Catch ex As Exception
                report.AppendLine("Process Info Error: " & ex.GetType().FullName & ": " & ex.Message)
            End Try

            report.AppendLine()
            report.AppendLine("CONFIGURATION")
            report.AppendLine(New String("-"c, 80))

            Try
                report.AppendLine("Active redink.ini Path: " & SafeValue(SharedMethods.GetActiveConfigFilePath(_context)))
            Catch ex As Exception
                report.AppendLine("Active redink.ini Path Error: " & ex.GetType().FullName & ": " & ex.Message)
            End Try

            report.AppendLine("INI_LogPath: " & SafeValue(_context.INI_LogPath))
            report.AppendLine("INI_UpdatePath: " & SafeValue(_context.INI_UpdatePath))
            report.AppendLine()

            report.AppendLine("KEY ASSEMBLIES")
            report.AppendLine(New String("-"c, 80))
            AppendAssemblyInfo(report, "ThisAddIn Assembly", Me.GetType().Assembly)
            AppendAssemblyInfo(report, "SharedLibrary Assembly", GetType(SharedMethods).Assembly)
            AppendAssemblyInfo(report, "Newtonsoft.Json / JToken Assembly", GetType(Newtonsoft.Json.Linq.JToken).Assembly)

            report.AppendLine()
            report.AppendLine("JTOKEN METHOD CHECK")
            report.AppendLine(New String("-"c, 80))

            Try
                Dim jTokenType As System.Type = GetType(Newtonsoft.Json.Linq.JToken)
                Dim formattingType As System.Type = GetType(Newtonsoft.Json.Formatting)
                Dim formattingMethod As System.Reflection.MethodInfo = jTokenType.GetMethod("ToString", New System.Type() {formattingType})

                report.AppendLine("JToken Type AssemblyQualifiedName: " & SafeValue(jTokenType.AssemblyQualifiedName))
                If formattingMethod Is Nothing Then
                    report.AppendLine("Reflection Lookup: JToken.ToString(Newtonsoft.Json.Formatting) = NOT FOUND")
                Else
                    report.AppendLine("Reflection Lookup: " & formattingMethod.ToString())
                End If
            Catch ex As Exception
                report.AppendLine("Reflection Lookup Error: " & ex.GetType().FullName & ": " & ex.Message)
            End Try

            report.AppendLine()
            report.AppendLine("RUNTIME PROBE")
            report.AppendLine(New String("-"c, 80))

            Try
                Dim probeToken As Newtonsoft.Json.Linq.JToken = Newtonsoft.Json.Linq.JToken.Parse("{""x"":1}")
                Dim probeText As String = probeToken.ToString(Newtonsoft.Json.Formatting.Indented)

                report.AppendLine("Probe Result: SUCCESS")
                report.AppendLine("Probe Output:")
                report.AppendLine(probeText)
            Catch ex As Exception
                report.AppendLine("Probe Result: FAILED")
                report.AppendLine("Exception Type: " & ex.GetType().FullName)
                report.AppendLine("Exception Message: " & ex.Message)
                report.AppendLine("Stack Trace:")
                report.AppendLine(If(ex.StackTrace, ""))
            End Try

            report.AppendLine()
            report.AppendLine("LOADED ASSEMBLIES OF INTEREST")
            report.AppendLine(New String("-"c, 80))

            Try
                For Each loadedAssembly As System.Reflection.Assembly In AppDomain.CurrentDomain.GetAssemblies()
                    Dim assemblyName As String = ""
                    Try
                        assemblyName = loadedAssembly.GetName().Name
                    Catch
                    End Try

                    If String.Equals(assemblyName, "Newtonsoft.Json", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(assemblyName, "SharedLibrary", StringComparison.OrdinalIgnoreCase) OrElse
                       String.Equals(assemblyName, Me.GetType().Assembly.GetName().Name, StringComparison.OrdinalIgnoreCase) Then

                        AppendAssemblyInfo(report, "Loaded Assembly", loadedAssembly)
                        report.AppendLine()
                    End If
                Next
            Catch ex As Exception
                report.AppendLine("Loaded Assembly Scan Error: " & ex.GetType().FullName & ": " & ex.Message)
            End Try

            report.AppendLine("APP.CONFIG NEWTONSOFT SNIPPET")
            report.AppendLine(New String("-"c, 80))
            AppendConfigSnippet(report, AppDomain.CurrentDomain.SetupInformation.ConfigurationFile, "Newtonsoft.Json")

            report.AppendLine()
            report.AppendLine("DONE")
            report.AppendLine(New String("="c, 80))

            System.IO.File.WriteAllText(outputPath, report.ToString(), New System.Text.UTF8Encoding(False))
        Catch
            ShowCustomMessageBox("Cannot generate RI DLL load diagnostics due to an error.")
            ' Intentionally silent: diagnostics must never break startup.
        End Try
    End Sub

    Private Shared Sub AppendAssemblyInfo(report As System.Text.StringBuilder, title As String, assemblyValue As System.Reflection.Assembly)
        report.AppendLine(title & ":")

        If assemblyValue Is Nothing Then
            report.AppendLine("  <nothing>")
            Return
        End If

        Try
            report.AppendLine("  FullName: " & SafeValue(assemblyValue.FullName))
        Catch ex As Exception
            report.AppendLine("  FullName Error: " & ex.GetType().FullName & ": " & ex.Message)
        End Try

        Try
            report.AppendLine("  Location: " & SafeValue(assemblyValue.Location))
        Catch ex As Exception
            report.AppendLine("  Location Error: " & ex.GetType().FullName & ": " & ex.Message)
        End Try

        Try
            report.AppendLine("  ImageRuntimeVersion: " & SafeValue(assemblyValue.ImageRuntimeVersion))
        Catch ex As Exception
            report.AppendLine("  ImageRuntimeVersion Error: " & ex.GetType().FullName & ": " & ex.Message)
        End Try

        Try
            Dim fileVersion As String = ""
            If Not String.IsNullOrWhiteSpace(assemblyValue.Location) AndAlso System.IO.File.Exists(assemblyValue.Location) Then
                fileVersion = System.Diagnostics.FileVersionInfo.GetVersionInfo(assemblyValue.Location).FileVersion
            End If
            report.AppendLine("  FileVersion: " & SafeValue(fileVersion))
        Catch ex As Exception
            report.AppendLine("  FileVersion Error: " & ex.GetType().FullName & ": " & ex.Message)
        End Try
    End Sub

    Private Shared Sub AppendConfigSnippet(report As System.Text.StringBuilder, configPath As String, searchText As String)
        Try
            If String.IsNullOrWhiteSpace(configPath) Then
                report.AppendLine("<no config path>")
                Return
            End If

            report.AppendLine("Config Path: " & configPath)

            If Not System.IO.File.Exists(configPath) Then
                report.AppendLine("<config file not found>")
                Return
            End If

            Dim lines As String() = System.IO.File.ReadAllLines(configPath)
            Dim found As Boolean = False

            For i As Integer = 0 To lines.Length - 1
                If lines(i).IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0 Then
                    found = True
                    Dim startIndex As Integer = Math.Max(0, i - 3)
                    Dim endIndex As Integer = Math.Min(lines.Length - 1, i + 3)

                    For j As Integer = startIndex To endIndex
                        report.AppendLine((j + 1).ToString(CultureInfo.InvariantCulture).PadLeft(5) & ": " & lines(j))
                    Next

                    Exit For
                End If
            Next

            If Not found Then
                report.AppendLine("<search text not found in config>")
            End If
        Catch ex As Exception
            report.AppendLine("Config Snippet Error: " & ex.GetType().FullName & ": " & ex.Message)
        End Try
    End Sub

    Private Shared Function SafeValue(value As String) As String
        If String.IsNullOrWhiteSpace(value) Then Return "<empty>"
        Return value
    End Function


End Class


' =================================================================================================
' Red Ink – Architectural Overview (for maintenance & security review)
'
' PURPOSE
'   Red Ink is a Word COM/VSTO Add-in providing AI-assisted authoring, search, review, redaction,
'   speech (TTS/STT), transcription and document transformation features. It orchestrates Word UI
'   events, user prompts, external AI services and local processing utilities.
'
' HIGH-LEVEL LAYERS
'   1. Host Integration (Word object model, events, task panes, context menus, Ribbon)
'   2. Command & Orchestration Layer (user actions → pipelines)
'   3. Processing & Transformation (text, markup, diff, formatting, RAG, grounding, redaction)
'   4. Speech & Transcription (TTS/STT handling, language model bridging)
'   5. Web / External Services (LLM calls, auxiliary web agent, update mechanism)
'   6. Helpers & Utilities (file I/O, Word-specific helpers, JSON templating, state management)
'
' CORE ENTRYPOINT
'   ThisAddIn.vb
'       - Startup sequencing: attaches Word events, then runs DelayedStartupTasks once a document
'         context exists (avoids premature automation calls).
'       - Initializes configuration via SharedLibrary (SharedMethods.Initialize / InitializeConfig).
'       - Registers context menus (AddContextMenu) & periodic update checks (UpdateHandler).
'       - Exposes LLM(), PostCorrection(), and other bridges to SharedLibrary for central API usage.
'       - Manages asynchronous UI thread marshaling (EnsureUIThread + mainThreadControl).
'       - Maintains numerous constants (trigger tokens, feature switches, language/model lists).
'       - Holds transient state fields used by runtime prompt interpolation.
'       - Performs COM automation exposure (RequestComAddInAutomationService → BridgeSubs).
'       - Security-sensitive areas: HTTP listener (StartupHttpListener/ShutdownHttpListener),
'         external endpoints (OpenAI, Google, Whisper, Vosk), dynamic JSON template formatting code.
'
' HOST / WORD INTEGRATION
'   Word (implicit)              – Microsoft.Office.Interop.Word objects consumed throughout.
'   ThisAddIn.WordHelpers.vb     – Word-range/document manipulations (selection, insertion, cleanup).
'   ThisAddIn.WordSearchHelper.vb– Specialized search routines (clauses, hidden prompts, regex etc.).
'   ThisAddIn.Menu.vb            – Context menu creation & teardown; maps triggers → command handlers.
'   ThisAddIn.PaneAndMerge.vb    – Task pane orchestration, pane content updates, merge/view logic.
'   ThisAddIn.Slides.vb          – Document → slide generation utilities.
'   ThisAddIn.Redactions.vb      – Sensitive content detection & redaction application.
'
' UI COMPONENTS
'   Ribbon1.vb                   – Ribbon callbacks (buttons → Commands layer). Multiple partials
'                                  may exist to segregate feature groups (ensure all are reviewed).
'   Form1.vb                     – Main dialog or configuration/interaction form (general UI).
'   DragDropForm.vb              – Drag-and-drop ingestion (files/clipboard objects → processing).
'   ThisAddIn.TextToSpeech.Form.vb – UI for TTS settings, voice/model selection, playback controls.
'
' VBA / AUTOMATION BRIDGE
'   VBA Helper (external .dotm)  – Provides macro functions (e.g., CheckAppHelper) to validate helper
'                                  version or offer legacy automation tasks. Security: ensure trusted
'                                  location + signed macros.
'   BridgeSubs.vb                – COM-visible automation surface (exposed via RequestComAddInAutomationService)
'                                  enabling external scripts/macros to trigger add-in commands safely.
'
' COMMAND LAYER (User Intent → Pipelines)
'   ThisAddIn.Commands.vb              – Central dispatcher for high-level actions.
'   ThisAddIn.Commands.Freestyle.vb    – Freeform / ad-hoc prompt execution and content generation.
'   ThisAddIn.TextToSpeech.Commands.vb – TTS/STT specific commands (start synthesis, transcription).
'
' SEARCH / CONTEXT / RETRIEVAL
'   ThisAddIn.ContextSearch.vb         – Multi-source contextual search (library, document chunks).
'   ThisAddIn.SearchGrounding.vb       – Grounding retrieved context before LLM calls.
'   ThisAddIn.FileRAG.vb               – File-based Retrieval-Augmented Generation (embeddings, chunking).
'   ThisAddIn.FindClause.vb            – Clause / structured legal text identification.
'   ThisAddIn.FindHiddenPrompt.vb      – Detection of hidden / obfuscated prompt injections (security).
'   ThisAddIn.DocCheck.vb              – Document diagnostics (integrity, version, compliance).
'
' PROCESSING / TRANSFORMATION
'   ThisAddIn.Processing.vb                    – Orchestrates multi-step pipelines (markup, diff, inject).
'   ThisAddIn.Processing.Comments.vb           – Bubble/comment extraction, formatting, pushback handling.
'   ThisAddIn.Processing.HTMLToWord.vb         – HTML → Word conversion (sanitization + formatting).
'   ThisAddIn.Processing.FormatSaveAndRestore.vb – Captures/restores formatting state around operations.
'   ThisAddIn.Processing.SearchGrounding.vb    – (See above) bridging search results into prompts.
'
' TEXT / DATA UTILITIES
'   ThisAddIn.FileHelpers.vb           – Safe file I/O, temp handling, path resolution, model file mgmt.
'   ThisAddIn.Helpers.vb               – Generic helper routines (parsing, mapping, template injection).
'   ThisAddIn.Properties.vb            – Configuration property wrappers (INI/env/registry abstraction).
'
' SPEECH / TRANSCRIPTION
'   ThisAddIn.TextToSpeech.vb          – Core TTS pipeline (Google, OpenAI, local fallback).
'   ThisAddIn.Transcriptor.vb          – STT/transcription (Whisper, Vosk, Google streaming).
'
' WEB / EXTERNAL INTERFACES
'   ThisAddIn.WebAgent.vb              – Local HTTP listener / agent for inter-process automation.
'   ThisAddIn.WebExtension.vb          – Bridges browser/extension requests to internal commands.
'   SpecialServices.vb (ThisAddIn.SpecialServices.vb)
'                                     – Ancillary external service tasks (e.g., update fetch, licensing).
'
' RESOURCES
'   Resources (resx / assets)          – Icons, localization strings, model metadata, license text.
'                                        Security: validate no embedded secrets; ensure license compliance.
'
' KEY DATA FLOWS
'   Word Events → ThisAddIn Startup → Menu/Ribbon & Pane Initialization → User Action →
'   Commands Layer → (Search/RAG/Processing) → LLM/TTS/STT/Web → Results Injection (markup, panes, doc).
'
' SECURITY / REVIEW HOTSPOTS
'   - External service calls (LLM/TTS/STT endpoints): validate secure transport (HTTPS / WSS).
'   - Dynamic code string (Code_JsonTemplateFormatter): ensure it is not compiled/executed unsafely.
'   - HTTP listener (WebAgent): restrict origin, sanitize inputs.
'   - Prompt injection & hidden content (FindHiddenPrompt.vb): confirm effective sanitization.
'   - File ingestion (DragDropForm, FileHelpers): path traversal & format validation.
'   - Macro bridge (VBA Helper + BridgeSubs): enforce version check & signed VBA project.
'   - Redaction logic: confirm irreversible removal when required.
'
'
' =================================================================================================
