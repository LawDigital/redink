' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' 10.1.2026
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
' Includes Google Speech V1 library and related API libraries in unchanged form; Copyright (c) 2024 Google LLC; licensed under the Apache 2.0 license (https://licenses.nuget.org/Apache-2.0) at https://github.com/googleapis/google-cloud-dotnet
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
' Includes also various Microsoft distributables and libraries copyrighted by Microsoft Corporation and available, among others, under the Microsoft EULA, the Visual Studio Community 2022 License, the Microsoft.Web.WebView2 License (for Microsoft.Web.WebView2, see license on https://www.nuget.org/packages/Microsoft.Web.WebView2/ and below) and the MIT License (including Microsoft.Bcl.*, Microsoft.Extensions.*, System.*, System.Security.*, System.CodeDom, DocumentFormat.OpenXml.*, Microsoft.ml.*, CommunityToolkit.HighPerformance licensed under MIT License) (https://licenses.nuget.org/MIT); Copyright (c) 2016- Microsoft Corp.

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

Partial Public Class ThisAddIn

    ' Hardcoded config values

    Public Shared Version As String = "V.100126" & SharedMethods.VersionQualifier

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
    Private Const BubbleCutText As String = " (" & ChrW(&H2702) & ")"
    Private Const SearchNextTrigger As String = "Next:"
    Private Const BoWTrigger As String = "(bow)"
    Private Const ChunkTrigger As String = "(iterate)"
    Private Const EmbedTrigger As String = "(embed)"
    Private Const RefreshTrigger As String = "(refresh)"
    Private Const ToolSelectionTrigger As String = "(sources)"  ' Trigger in OtherPrompt to re-select tools for tooling-enabled models.
    Public Const ToolFriendlyName As String = "Sources"  ' How to refer to tools (e.g., sources) towards the user

    Private Const MaxFilibuster As Integer = 10000 ' Maximum number of words for filibuster mode 
    Private Const ArgueAgainstDefault As Integer = 50 ' Number of words to propose for Argue Against

    Private Const RegexSeparator1 As String = "|||"  ' Set also in SharedLibrary
    Private Const RegexSeparator2 As String = "§§§"  ' Set also in SharedLibrary 
    Private Const RIMenu = AN
    Private Const OldRIMenu = AN & " " & ChrW(&HD83D) & ChrW(&HDC09)
    Private Const MinHelperVersion = 1 ' Minimum version of the helper file that is required

    Public Const IgnoreMarkups As Boolean = False ' Whether to ignore markups in the text when doing a search

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
    {"alloy", "Female: Versatile and balanced"},
    {"ash", "Male: Clear and precise"},
    {"ballad", "Male: Melodic and smooth"},
    {"coral", "Female: Warm and friendly"},
    {"echo", "Male: Warm and natural"},
    {"fable", "Male: Engaging storyteller"},
    {"nova", "Female: Bright and energetic"},
    {"onyx", "Male: Deep and authoritative"},
    {"sage", "Male: Calm and thoughtful"},
    {"shimmer", "Female: Clear and expressive"},
    {"verse", "Male: Versatile and expressive"}
}

    Private Const TTS_OpenAI_Model = "gpt-4o-mini-tts"

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

    Public Shared GoogleSTT_Desc As String = "Google STT V1 (run in EU)"
    Public Shared STTEndpoint As String = "eu-speech.googleapis.com"


    Public Shared OpenAISTTModel As String = "gpt-4o-realtime-preview"
    Public Shared OpenAISTT_Desc As String = $"OpenAI Streaming"
    Public Shared STTEndpoint_OpenAI As String = $"wss://api.openai.com/v1/realtime?model=gpt-4o-realtime-preview"

    Public Shared GoogleSTTsupportedLanguages As String() = {
    "en-US", "de-DE",
    "de-AT", "de-CH", "es-AR", "es-BO", "es-CL", "es-CO", "es-CR", "es-DO", "es-EC", "es-ES", "es-GT",
    "es-HN", "es-MX", "es-NI", "es-PA", "es-PE", "es-PR", "es-PY", "es-SV", "es-UY", "es-VE",
    "fr-BE", "fr-CA", "fr-CH", "fr-FR", "it-CH", "it-IT", "nl-BE", "nl-NL",
    "af-ZA", "am-ET", "ar-BH", "ar-DZ", "ar-EG", "ar-IQ", "ar-IL", "ar-JO", "ar-KW", "ar-LB", "ar-MA",
    "ar-MR", "ar-OM", "ar-PS", "ar-QA", "ar-SA", "ar-SY", "ar-TN", "ar-AE", "ar-YE", "az-AZ", "bg-BG",
    "bn-BD", "bn-IN", "bs-BA", "ca-ES", "cmn-Hans-CN", "cmn-Hans-HK", "cmn-Hant-TW", "cs-CZ",
    "da-DK", "el-GR", "en-AU", "en-CA", "en-GH", "en-HK", "en-IE", "en-IN", "en-KE", "en-NG",
    "en-NZ", "en-PH", "en-PK", "en-SG", "en-TZ", "en-ZA", "et-EE", "eu-ES", "fa-IR", "fi-FI",
    "fil-PH", "gl-ES", "gu-IN", "hi-IN", "hr-HR", "hu-HU", "hy-AM", "id-ID", "is-IS", "iw-IL",
    "ja-JP", "jv-ID", "ka-GE", "kk-KZ", "km-KH", "kn-IN", "ko-KR", "lo-LA", "lt-LT", "lv-LV",
    "ml-IN", "mn-MN", "mr-IN", "ms-MY", "my-MM", "ne-NP", "no-NO", "pa-Guru-IN", "pl-PL",
    "pt-BR", "pt-PT", "ro-RO", "rw-RW", "si-LK", "sk-SK", "sl-SI", "sr-RS", "ss-Latn-ZA", "st-ZA",
    "su-ID", "sv-SE", "sw-KE", "sw-TZ", "ta-IN", "ta-LK", "ta-MY", "ta-SG", "te-IN", "th-TH",
    "tn-Latn-ZA", "tr-TR", "uk-UA", "ur-IN", "ur-PK", "uz-UZ", "ve-ZA", "vi-VN", "xh-ZA",
    "yue-Hant-HK", "zu-ZA"
        }

    ' Tooling

    Public Const ToolingLog_AutoCloseDefaultSeconds As Integer = 20

    Public Const InternalToolSuffix As String = " (internal)"  ' Suffix displayed for the internal web tool in selection dialogs.

    Public Const InternalWebToolName As String = "web_content_retriever"
    Public Const InternalWebToolDescription As String =
        "Retrieves readable text from one or more web pages. Use this tool when you need to access the content behind a URL instead of relying on summaries or excerpts."

    Public Const InternalWebToolDefinition As String =
        "{""name"":""web_content_retriever"",""description"":""Fetches and returns readable text from one or more web URLs."",""parameters"":{""type"":""object"",""properties"":{""urls"":{""type"":""array"",""items"":{""type"":""string""},""description"":""One or more absolute URLs to fetch (preferred).""},""url"":{""type"":""string"",""description"":""Single absolute URL to fetch (alternative to urls).""}}}}" ' Note: do not require urls; code validates at runtime.

    Public Const InternalWebToolInstructionsPrompt As String =
        "web_content_retriever: Fetches readable text from web pages. " &
        "Call this tool when you need the actual page content behind a link. " &
        "Provide either urls (array of strings) or url (single string). " &
        "Return value is plain text content for each URL (or an error per URL if retrieval fails)."

    Public Shared SelectedToolNames As New List(Of String)()   ' Persisted list of selected tool names for tooling sessions.


    ' Declare variables publicly so that InterpolateAtRuntime can access them; case-sensitive

    Public TranslateLanguage As String
    Public ShortenLength As Double
    Public FilibusterLength As Integer
    Public SummaryLength As Integer
    Public OtherPrompt As String = ""
    Public OutputLanguage As String = ""
    Public MaxToolIterations As Integer = 10
    Public InsertDocs As String = ""
    Public MyStyleInsert As String = ""
    Public FormatInstruction As String = ""
    Public SearchTerms As String
    Public SearchContext As String
    Public CurrentDate As String = "(Current Date: " & DateTime.Now.ToString("dd-MMM-yyyy", CultureInfo.GetCultureInfo("en-US")) & ")"
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


    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ' Necessary for Update Handler to work correctly

        ' 1) Force the creation of the Control's handle on the Office UI thread
        Dim dummy = mainThreadControl.Handle

        ' 2) Capture synchronization context & scheduler after handle exists
        _uiContext = SynchronizationContext.Current
        If _uiContext Is Nothing Then
            _uiContext = New WindowsFormsSynchronizationContext()
            SynchronizationContext.SetSynchronizationContext(_uiContext)
        End If
        _uiScheduler = TaskScheduler.FromCurrentSynchronizationContext()

        ' 3) Give that Control to the UpdateHandler so it can Invoke on it
        UpdateHandler.MainControl = mainThreadControl

        ' 4) Capture the host window’s HWND (Word / Excel / Outlook)
        Dim hwnd As IntPtr
        Dim progId = Me.Application.GetType().Name.ToLowerInvariant()
        If progId.Contains("word") OrElse progId.Contains("excel") Then
            hwnd = New IntPtr(CInt(Me.Application.Hwnd))
        Else
            hwnd = FindWindow("rctrl_renwnd32", Nothing)
        End If
        UpdateHandler.HostHandle = hwnd

        ' Other tasks that need to be done at startup

        SharedMethods.Initialize(Me.CustomTaskPanes)

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
            StartupHttpListener()
        Catch ex As System.Exception
            ' Handle exceptions gracefully.
        End Try
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveOldContextMenu()
        ShutdownHttpListener()
    End Sub

    Public Sub InitializeAddInFeatures()
        InitializeConfig(True, True)
        AddContextMenu()
        UpdateHandler.PeriodicCheckForUpdates(INI_UpdateCheckInterval, RDV, INI_UpdatePath, _context)
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
    Public Shared Async Function LLM(ByVal promptSystem As String, ByVal promptUser As String, Optional ByVal Model As String = "", Optional ByVal Temperature As String = "", Optional ByVal Timeout As Long = 0, Optional ByVal UseSecondAPI As Boolean = False, Optional ByVal Hidesplash As Boolean = False, Optional ByVal AddUserPrompt As String = "", Optional ByVal FileObject As String = "", Optional ByVal ToolExecution As Boolean = False) As Task(Of String)
        Dim Response = Await SharedMethods.LLM(_context, promptSystem, promptUser, Model, Temperature, Timeout, UseSecondAPI, Hidesplash, AddUserPrompt, FileObject, ToolExecution:=ToolExecution)
        Await EnsureUIThread().ConfigureAwait(False)
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