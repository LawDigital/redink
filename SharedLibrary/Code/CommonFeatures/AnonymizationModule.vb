' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: AnonymizationModule.vb
' Purpose:
'   Provides text anonymization and re-identification based on an external
'   configuration file (`redink-anon.txt`) and/or user input. Detected entities
'   are replaced with placeholders and recorded in-memory so placeholders can
'   later be re-identified during the current session.
'
' Architecture / How it works:
'   - Settings Lookup:
'       `LoadAnonSettingsForModel` reads `redink-anon.txt` and returns the
'       "mode; type" tuple for a given model name. Model-specific sections
'       override `[All]`.
'   - Mode / Type Parsing:
'       `GetModeFromSettings` and `GetTypeFromSettings` parse the settings tuple.
'   - Pattern Compilation:
'       Patterns are compiled either:
'         * from the file (`CompilePatternsForModel`, `BuildDefaultPromptFromFile`)
'         * or from user input (`BuildPatternInfosFromRawInput`)
'       Supported entries include:
'         * `Regex:` lines (raw regex pattern, optional `{{prefix}}`)
'         * literal tokens (regex-escaped)
'         * wildcard tokens containing `*` (the `*` is converted into `[\p{L}\p{N}-]*`)
'         * comma-separated tokens, where quoted strings ("...") are treated as one token
'   - Replacement Strategy:
'       `AnonymizeText` repeatedly finds the earliest match among all patterns,
'       replaces it with a placeholder built from a prefix + GroupID + sub-index,
'       and stores the placeholder -> original mapping in `EntitiesMappings`.
'       For the same GroupID, repeated occurrences of the same matched value reuse
'       the same sub-index.
'   - Review UI:
'       For modes that require review (`show`, `askshow`) the anonymized text is
'       displayed for editing via `ShowCustomWindow`.
'   - Re-identification:
'       `ReidentifyText` replaces placeholders using the current in-memory mappings.
'
' Configuration file structure (`redink-anon.txt` at `AnonFilepath`):
'
'   ; Comment lines start with semicolon
'
'   [All]
'   Anon = mode; type
'
'   [ModelName1, ModelName2]
'   Anon = mode; type
'   Regex:regexcode
'   ENTITY1
'   ENTITY2*{{placeholder}}
'   ENTITY3, EnTITY4, ENTITY5
'
' Sections:
'   [All] applies to any model. Subsequent lines until next [Section] apply to All.
'   [ModelName, OtherModel] applies only to those models. In case of conflict,
'   model-specific overrides [All].
'
' Lines under a section:
'   Anon = mode; type
'     - mode = none, silent, ask, askshow, show
'     - type = 0 (none), 1 (user prompt with last prompt default), 2 (user prompt empty),
'              3 (file-based only), 4 (user prompt with file-based default)
'   Regex:pattern      (regular expression pattern; may include {{prefix}} to override placeholder)
'   ENTITY literal     (exact match, escaped for regex)
'   WILDCARD*          (wildcard '*' converts to "[\p{L}\p{N}-]*")
'   Multiple entities can be comma-separated on one line; quoted strings ("multi word")
'   are treated as single terms.
'
' Placeholder format: <prefix_GGGG_SSS>
'   - prefix: default "redacted" or custom via {{prefix}}
'   - GGGG: 4-digit GroupID (unique per pattern)
'   - SSS:  sub-index (starts at 1 for first distinct match of that pattern,
'          increments for subsequent distinct matches)
'
' Modes:
'   "none"    = No anonymization.
'   "silent"  = Anonymize automatically without prompts or previews.
'   "ask"     = Prompt Yes/No. If Yes, anonymize silently.
'   "askshow" = Prompt Yes/No. If Yes, anonymize then show for editing.
'   "show"    = Always anonymize, then show for editing.
'
' Types:
'   0 = No anonymization.
'   1 = Prompt user; default = last-used prompt (My.Settings.LastAnonPrompt).
'   2 = Prompt user; default = empty.
'   3 = Use only patterns from file; no UI prompt.
'   4 = Prompt user; default = literals/wildcards from file.
' =============================================================================


Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary

    ''' <summary>
    ''' Provides anonymization and re-identification helpers based on a configuration file and/or user input.
    ''' </summary>
    Public Module AnonymizationModule

        ''' <summary>
        ''' Path to the anonymization configuration file on the Desktop.
        ''' </summary>
        Public AnonFilepath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), AnonFile)

        ''' <summary>
        ''' Default placeholder prefix used when no custom <c>{{prefix}}</c> marker is provided.
        ''' </summary>
        Private Const DEFAULT_PLACEHOLDER As String = AnonPlaceholder

        ''' <summary>
        ''' In-memory mapping of generated placeholders to original matched strings for the current session.
        ''' </summary>
        Private EntitiesMappings As New List(Of KeyValuePair(Of String, String))

        ''' <summary>
        ''' Holds a compiled regex and the metadata required to generate placeholders for this pattern.
        ''' </summary>
        Private Class PatternInfo

            ''' <summary>
            ''' Compiled regex used to detect matches in the working text.
            ''' </summary>
            Public Property RegexPattern As Regex

            ''' <summary>
            ''' Placeholder prefix for this pattern (default or overridden by <c>{{prefix}}</c>).
            ''' </summary>
            Public Property Prefix As String

            ''' <summary>
            ''' 1-based identifier assigned in compilation order; used as placeholder GroupID (GGGG).
            ''' </summary>
            Public Property GroupID As Integer

            ''' <summary>
            ''' Initializes a new <see cref="PatternInfo"/>.
            ''' </summary>
            Public Sub New(rx As Regex, prefix As String, groupID As Integer)
                Me.RegexPattern = rx
                Me.Prefix = prefix
                Me.GroupID = groupID
            End Sub
        End Class

        ' ------------------------------------------------------------------------
        ' 1. LoadAnonSettingsForModel(modelName) As String
        '    Reads redink-anon.txt and returns the "mode; type" for the given model.
        '    Searches [All] and [ModelName]; model-specific overrides [All].
        '    Returns empty string if no setting found or on error.
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Reads <c>redink-anon.txt</c> and returns the <c>"mode; type"</c> setting for a given model name.
        ''' </summary>
        ''' <param name="modelName">Model name used to match a section header (e.g., <c>[ModelA, ModelB]</c>).</param>
        ''' <returns>
        ''' The raw settings string (<c>"mode; type"</c>), an empty string if no setting was found,
        ''' or an empty string when the file does not exist.
        ''' </returns>
        Public Function LoadAnonSettingsForModel(ByVal modelName As String) As String
            Dim allSetting As String = String.Empty
            Dim modelSetting As String = String.Empty

            Try
                If Not File.Exists(AnonFilepath) Then
                    Return String.Empty
                End If

                Dim lines As String() = File.ReadAllLines(AnonFilepath)
                Dim currentSection As String = String.Empty
                Dim isAllSection As Boolean = False
                Dim isModelSection As Boolean = False

                For Each rawLine As String In lines
                    Dim line As String = rawLine.Trim()
                    If line.StartsWith(";") OrElse String.IsNullOrEmpty(line) Then
                        Continue For
                    End If

                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        currentSection = line.Substring(1, line.Length - 2).Trim()
                        isAllSection = String.Equals(currentSection, "All", StringComparison.OrdinalIgnoreCase)

                        If Not isAllSection Then
                            Dim modelTokens As String() = currentSection.Split(","c)
                            Dim found As Boolean = False
                            For Each tok In modelTokens
                                If String.Equals(tok.Trim(), modelName, StringComparison.OrdinalIgnoreCase) Then
                                    found = True
                                    Exit For
                                End If
                            Next
                            isModelSection = found
                        Else
                            isModelSection = False
                        End If

                        Continue For
                    End If

                    If line.StartsWith("Anon", StringComparison.OrdinalIgnoreCase) Then
                        Dim parts() As String = line.Split(New Char() {"="c}, 2)
                        If parts.Length = 2 Then
                            Dim valuePart As String = parts(1).Trim()
                            If isAllSection Then
                                allSetting = valuePart
                            ElseIf isModelSection Then
                                modelSetting = valuePart
                            End If
                        End If
                    End If
                Next

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error loading anonymization settings: {ex.Message}")
                Return String.Empty
            End Try

            ' Model-specific takes precedence.
            If Not String.IsNullOrWhiteSpace(modelSetting) Then
                Return modelSetting
            End If
            Return allSetting
        End Function

        ' ------------------------------------------------------------------------
        ' 2. GetModeFromSettings(settingsString) As String
        '    Splits "mode; type" and returns mode in lowercase, or empty if invalid.
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts the anonymization mode (the part before the first semicolon) from a settings string.
        ''' </summary>
        ''' <param name="settingsString">The settings string in the form <c>"mode; type"</c>.</param>
        ''' <returns>The lowercased mode; otherwise an empty string if not present.</returns>
        Public Function GetModeFromSettings(ByVal settingsString As String) As String
            Try
                If String.IsNullOrWhiteSpace(settingsString) Then
                    Return String.Empty
                End If
                Dim parts() As String = settingsString.Split(";"c)
                If parts.Length >= 1 Then
                    Return parts(0).Trim().ToLowerInvariant()
                End If
                Return String.Empty
            Catch ex As System.Exception
                ShowCustomMessageBox($"Error extracting mode: {ex.Message}")
                Return String.Empty
            End Try
        End Function

        ' ------------------------------------------------------------------------
        ' 3. GetTypeFromSettings(settingsString) As Integer
        '    Splits "mode; type" and returns type as integer, or 0 if invalid.
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Extracts the anonymization type (the integer after the semicolon) from a settings string.
        ''' </summary>
        ''' <param name="settingsString">The settings string in the form <c>"mode; type"</c>.</param>
        ''' <returns>The parsed type value; <c>0</c> if missing or invalid.</returns>
        Public Function GetTypeFromSettings(ByVal settingsString As String) As Integer
            Try
                If String.IsNullOrWhiteSpace(settingsString) Then
                    Return 0
                End If
                Dim parts() As String = settingsString.Split(";"c)
                If parts.Length >= 2 Then
                    Dim typePart As String = parts(1).Trim()
                    Dim result As Integer = 0
                    If Integer.TryParse(typePart, result) Then
                        Return result
                    End If
                End If
                Return 0
            Catch ex As System.Exception
                ShowCustomMessageBox($"Error extracting type: {ex.Message}")
                Return 0
            End Try
        End Function

        ' ------------------------------------------------------------------------
        ' 4. AnonymizeText(inputText, modelName, mode, typeValue) As String
        '    Performs anonymization based on mode and type for the specified model.
        '    Returns anonymized text or original text on error or "no anonymization".
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Replaces occurrences of configured/user-provided entities in <paramref name="inputText"/> with placeholders.
        ''' </summary>
        ''' <param name="inputText">The input text to anonymize.</param>
        ''' <param name="modelName">Model name used for retrieving file-based settings and patterns.</param>
        ''' <param name="mode">Anonymization mode (<c>none</c>, <c>silent</c>, <c>ask</c>, <c>askshow</c>, <c>show</c>).</param>
        ''' <param name="typeValue">Anonymization type (0..4) controlling whether patterns come from file and/or user prompt.</param>
        ''' <returns>
        ''' The anonymized text; the original text if no anonymization is requested; or an empty string on user cancel paths.
        ''' </returns>
        Public Function AnonymizeText(ByVal inputText As String,
                          ByVal modelName As String,
                          ByVal mode As String,
                          ByVal typeValue As Integer) As String

            Dim result As String = inputText

            Try
                ' 1) If no anonymization is requested:
                If String.IsNullOrEmpty(mode) OrElse mode = "none" OrElse typeValue = 0 Then
                    Return inputText
                End If

                ' 2) For "ask" or "askshow" prompt the user:
                If mode = "ask" OrElse mode = "askshow" Then
                    Dim promptText As String = "Do you want to anonymize?"
                    If mode = "askshow" Then
                        promptText = "Do you want to anonymize and see the text?"
                    End If

                    ' ShowCustomYesNoBox returns: 1 = Yes, 0 = No
                    Dim choice As Integer = ShowCustomYesNoBox(promptText, "Yes", "No", $"{AN} Anonymization")
                    If choice <> 1 Then
                        Return inputText
                    End If
                    ' Continue with anonymization
                End If

                ' 3) Build the pattern list (from file and/or prompt):
                Dim patternsList As New List(Of PatternInfo)()

                If typeValue = 3 Then
                    patternsList = CompilePatternsForModel(modelName)
                    If patternsList.Count = 0 AndAlso mode <> "silent" Then
                        ShowCustomMessageBox("No patterns found in file or file missing for type = 3.")
                        Return inputText
                    End If

                ElseIf typeValue = 4 Then
                    Dim defaultPrompt As String = BuildDefaultPromptFromFile(modelName)
                    Dim promptResponse As String = ShowCustomInputBox(
            $"Enter entities to anonymize (comma-separated); you can use wildcards and ""...""; default comes from your file '{AnonFilepath}':",
            $"{AN} Anonymization", False, defaultPrompt)

                    If promptResponse Is Nothing Then
                        Return inputText
                    End If
                    If promptResponse = "esc" Then
                        Return ""
                    End If
                    If String.IsNullOrWhiteSpace(promptResponse) Then
                        Return inputText
                    End If

                    patternsList = BuildPatternInfosFromRawInput(promptResponse)

                ElseIf typeValue = 1 OrElse typeValue = 2 Then
                    Dim defaultPrompt As String = String.Empty
                    If typeValue = 1 Then
                        defaultPrompt = If(My.Settings.LastAnonPrompt, String.Empty)
                    End If

                    Dim promptResponse As String = ShowCustomInputBox(
            $"Enter entities to anonymize (comma-separated); you can use wildcards and ""..."":",
            $"{AN} Anonymization", False, defaultPrompt)

                    If promptResponse Is Nothing Then
                        Return inputText
                    End If
                    If promptResponse = "esc" Then
                        Return ""
                    End If
                    If String.IsNullOrWhiteSpace(promptResponse) Then
                        Return inputText
                    End If

                    If typeValue = 1 Then
                        Try
                            My.Settings.LastAnonPrompt = promptResponse
                            My.Settings.Save()
                        Catch setEx As System.Exception
                            ShowCustomMessageBox($"Error saving settings: {setEx.Message}")
                        End Try
                    End If

                    patternsList = BuildPatternInfosFromRawInput(promptResponse)

                Else
                    Return inputText
                End If

                ' 4) Anonymization loop: always replace the earliest next match across all patterns.
                EntitiesMappings.Clear()
                Dim workingText As String = result

                ' For each GroupID:
                '  - keep a counter for new Sub-Indices
                '  - keep a dictionary mapping each matched value to its Sub-Index
                Dim groupSubCounters As New Dictionary(Of Integer, Integer)()
                Dim groupValueToIndex As New Dictionary(Of Integer, Dictionary(Of String, Integer))()

                For Each pi In patternsList
                    groupSubCounters(pi.GroupID) = 0
                    groupValueToIndex(pi.GroupID) = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
                Next

                While True
                    Dim earliestMatch As System.Text.RegularExpressions.Match = Nothing
                    Dim matchPatternInfo As PatternInfo = Nothing

                    ' Search across all patterns for the match with lowest index:
                    For Each pi In patternsList
                        Dim m As System.Text.RegularExpressions.Match = pi.RegexPattern.Match(workingText)
                        If m.Success Then
                            If earliestMatch Is Nothing OrElse m.Index < earliestMatch.Index Then
                                earliestMatch = m
                                matchPatternInfo = pi
                            End If
                        End If
                    Next

                    If earliestMatch Is Nothing OrElse matchPatternInfo Is Nothing Then
                        Exit While
                    End If

                    Dim grpID As Integer = matchPatternInfo.GroupID
                    Dim matchedValue As String = earliestMatch.Value

                    Dim subIndex As Integer
                    Dim placeholdersForGroup As Dictionary(Of String, Integer) = groupValueToIndex(grpID)

                    If placeholdersForGroup.ContainsKey(matchedValue) Then
                        ' If this exact matched value has already been seen for this group, reuse the same Sub-Index.
                        subIndex = placeholdersForGroup(matchedValue)
                    Else
                        ' New matched value for this group: increment Sub-Index counter.
                        Dim nextSub As Integer = groupSubCounters(grpID) + 1
                        groupSubCounters(grpID) = nextSub
                        subIndex = nextSub
                        placeholdersForGroup(matchedValue) = subIndex

                        ' Only on first occurrence of a new matched value in this group: store mapping.
                        Dim newPlaceholder As String = AnonPrefix & $"{matchPatternInfo.Prefix}_{grpID.ToString("D4")}_{subIndex}" & AnonSuffix
                        EntitiesMappings.Add(New KeyValuePair(Of String, String)(newPlaceholder, matchedValue))
                    End If

                    ' Build placeholder string:
                    Dim placeholder As String = AnonPrefix & $"{matchPatternInfo.Prefix}_{grpID.ToString("D4")}_{subIndex}" & AnonSuffix

                    ' Rebuild text:
                    Dim before As String = workingText.Substring(0, earliestMatch.Index)
                    Dim after As String = workingText.Substring(earliestMatch.Index + earliestMatch.Length)
                    workingText = before & placeholder & after

                End While

                result = workingText

                ' 5) For "show" or "askshow", show anonymized text for review/editing:
                If mode = "show" OrElse mode = "askshow" Then

                    'Debug.WriteLine(ExportEntitiesMappings)

                    Dim editedResponse As String = ShowCustomWindow(
            "Review your anonymized text. You may edit it before having it processed:",
            result,
            "You can choose to go on with the original text or your edits. Do not remove formatting code and do not change placeholders. Also avoid adding or removing lines, as this may distort the formatting of the results.",
            $"{AN} Anonymization", True, False)

                    If editedResponse Is Nothing OrElse editedResponse = "esc" OrElse String.IsNullOrWhiteSpace(editedResponse) Then
                        Return ""
                    End If

                    result = editedResponse
                End If

                Return result

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error during AnonymizeText: {ex.Message}")
                Return inputText
            End Try
        End Function



        ' ------------------------------------------------------------------------
        ' 5. ReidentifyText(inputText) As String
        '    Replaces placeholders in inputText with original entities from EntitiesMappings.
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Replaces placeholders in <paramref name="inputText"/> with original values from the current in-memory mapping.
        ''' </summary>
        ''' <param name="inputText">Text which may contain placeholders created by <see cref="AnonymizeText"/>.</param>
        ''' <returns>The text with known placeholders replaced by their original values.</returns>
        Public Function ReidentifyText(ByVal inputText As String) As String
            Try
                Dim output As String = inputText
                For Each kvp In EntitiesMappings
                    output = output.Replace(kvp.Key, kvp.Value)
                Next
                Return output
            Catch ex As System.Exception
                ShowCustomMessageBox($"Error during ReIdentifyText: {ex.Message}")
                Return inputText
            End Try
        End Function

        ' ------------------------------------------------------------------------
        ' 6. ExportEntitiesMappings() As String
        '    Returns the EntitiesMappings as a multi-line text:
        '      [prefix_0001_1]: OriginalEntity1
        '      [prefix_0002_1]: OriginalEntity2
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Exports the current placeholder-to-original mappings as a multi-line string.
        ''' </summary>
        ''' <returns>A multi-line mapping representation in the form <c>{placeholder}: {original}</c>.</returns>
        Public Function ExportEntitiesMappings() As String
            Try
                Dim sb As New StringBuilder()
                For Each kvp In EntitiesMappings
                    sb.AppendLine($"{kvp.Key}: {kvp.Value}")
                Next
                Return sb.ToString()
            Catch ex As System.Exception
                ShowCustomMessageBox($"Error exporting entity mappings: {ex.Message}")
                Return String.Empty
            End Try
        End Function


        ' ------------------------------------------------------------------------
        ' Helper: BuildPatternInfosFromRawInput(rawInput) As List(Of PatternInfo)
        '   Parses comma-separated tokens, detects "{{prefix}}", quotes and wildcards.
        '   Converts "*" to "[\p{L}\p{N}-]*".
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Parses a comma-separated user input string into a list of compiled patterns.
        ''' </summary>
        ''' <param name="rawInput">
        ''' Comma-separated tokens; quoted tokens (<c>"..."</c>) are treated as one token; tokens may contain
        ''' wildcards (<c>*</c>) and an optional prefix marker (<c>{{prefix}}</c>).
        ''' </param>
        ''' <returns>A list of compiled patterns with assigned GroupIDs, in token order.</returns>
        Private Function BuildPatternInfosFromRawInput(ByVal rawInput As String) As List(Of PatternInfo)
            Dim patternInfos As New List(Of PatternInfo)()

            Try
                Dim tokens As New List(Of String)()
                Dim current As New StringBuilder()
                Dim inQuotes As Boolean = False

                For i As Integer = 0 To rawInput.Length - 1
                    Dim ch As Char = rawInput(i)
                    If ch = """"c Then
                        inQuotes = Not inQuotes
                        current.Append(ch)
                    ElseIf ch = ","c AndAlso Not inQuotes Then
                        tokens.Add(current.ToString().Trim())
                        current.Clear()
                    Else
                        current.Append(ch)
                    End If
                Next
                If current.Length > 0 Then
                    tokens.Add(current.ToString().Trim())
                End If

                Dim groupIDCounter As Integer = 0

                For Each rawTok As String In tokens
                    Dim tok As String = rawTok.Trim()
                    If String.IsNullOrEmpty(tok) Then
                        Continue For
                    End If

                    ' Detect custom prefix marker "{{prefix}}"
                    Dim prefix As String = DEFAULT_PLACEHOLDER
                    Dim tokenWithoutMarker As String = tok
                    Dim markerStart As Integer = tok.IndexOf("{{")
                    Dim markerEnd As Integer = tok.IndexOf("}}")
                    If markerStart >= 0 AndAlso markerEnd > markerStart Then
                        Dim between As String = tok.Substring(markerStart + 2, markerEnd - markerStart - 2).Trim()
                        If Not String.IsNullOrEmpty(between) Then
                            prefix = between
                        End If
                        tokenWithoutMarker = tok.Remove(markerStart, (markerEnd + 2) - markerStart).Trim()
                    End If

                    ' Determine regex pattern from tokenWithoutMarker (wildcard → "[\p{L}\p{N}-]*")
                    Dim patternText As String = String.Empty
                    If tokenWithoutMarker.StartsWith("""") AndAlso tokenWithoutMarker.EndsWith("""") AndAlso tokenWithoutMarker.Length >= 2 Then
                        Dim inner As String = tokenWithoutMarker.Substring(1, tokenWithoutMarker.Length - 2)
                        patternText = Regex.Escape(inner)
                    ElseIf tokenWithoutMarker.Contains("*") Then
                        Dim sbPat As New StringBuilder()
                        For Each c As Char In tokenWithoutMarker
                            If c = "*"c Then
                                sbPat.Append("[\p{L}\p{N}-]*")  ' wildcard matches letters/numbers/hyphen only
                            Else
                                sbPat.Append(Regex.Escape(c.ToString()))
                            End If
                        Next
                        patternText = sbPat.ToString()
                    Else
                        patternText = Regex.Escape(tokenWithoutMarker)
                    End If

                    Try
                        Dim rx As New Regex(patternText, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                        groupIDCounter += 1
                        patternInfos.Add(New PatternInfo(rx, prefix, groupIDCounter))
                    Catch rgEx As System.Exception
                        ShowCustomMessageBox($"Invalid pattern '{patternText}': {rgEx.Message}")
                    End Try
                Next

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error building patterns from input: {ex.Message}")
            End Try

            Return patternInfos
        End Function

        ' ------------------------------------------------------------------------
        ' Helper: CompilePatternsForModel(modelName) As List(Of PatternInfo)
        '   Reads file under [All] and [ModelName], processes "Regex:" lines
        '   and literal/wildcard lines. Converts "*" to "[\p{L}\p{N}-]*".
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Reads <c>redink-anon.txt</c> and compiles patterns from matching sections (<c>[All]</c> and model sections).
        ''' </summary>
        ''' <param name="modelName">Model name used to determine whether a section applies.</param>
        ''' <returns>A list of compiled patterns; empty if file is missing or no matching entity lines exist.</returns>
        Private Function CompilePatternsForModel(ByVal modelName As String) As List(Of PatternInfo)
            Dim patternInfos As New List(Of PatternInfo)()

            Try
                If Not File.Exists(AnonFilepath) Then
                    Return patternInfos
                End If

                Dim lines As String() = File.ReadAllLines(AnonFilepath)
                Dim currentSection As String = String.Empty
                Dim isAllSection As Boolean = False
                Dim isModelSection As Boolean = False

                Dim entityLines As New List(Of String)()

                For Each rawLine As String In lines
                    Dim line As String = rawLine.Trim()
                    If line.StartsWith(";") OrElse String.IsNullOrEmpty(line) Then
                        Continue For
                    End If

                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        currentSection = line.Substring(1, line.Length - 2).Trim()
                        isAllSection = String.Equals(currentSection, "All", StringComparison.OrdinalIgnoreCase)

                        If Not isAllSection Then
                            Dim modelTokens As String() = currentSection.Split(","c)
                            Dim found As Boolean = False
                            For Each tok In modelTokens
                                If String.Equals(tok.Trim(), modelName, StringComparison.OrdinalIgnoreCase) Then
                                    found = True
                                    Exit For
                                End If
                            Next
                            isModelSection = found
                        Else
                            isModelSection = False
                        End If

                        Continue For
                    End If

                    If line.StartsWith("Anon", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If isAllSection OrElse isModelSection Then
                        entityLines.Add(line)
                    End If
                Next

                Dim groupIDCounter As Integer = 0

                For Each item As String In entityLines
                    If item.StartsWith("Regex:", StringComparison.OrdinalIgnoreCase) Then
                        Dim remainder As String = item.Substring("Regex:".Length).Trim()
                        Dim prefix As String = DEFAULT_PLACEHOLDER
                        Dim patternRaw As String = remainder

                        Dim markerStart As Integer = remainder.IndexOf("{{")
                        Dim markerEnd As Integer = remainder.IndexOf("}}")
                        If markerStart >= 0 AndAlso markerEnd > markerStart Then
                            Dim between As String = remainder.Substring(markerStart + 2, markerEnd - markerStart - 2).Trim()
                            If Not String.IsNullOrEmpty(between) Then
                                prefix = between
                            End If
                            patternRaw = remainder.Remove(markerStart, (markerEnd + 2) - markerStart).Trim()
                        End If

                        Try
                            Dim rx As New Regex(patternRaw, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                            groupIDCounter += 1
                            patternInfos.Add(New PatternInfo(rx, prefix, groupIDCounter))
                        Catch rgEx As System.Exception
                            ShowCustomMessageBox($"Invalid regex '{patternRaw}': {rgEx.Message}")
                        End Try

                    Else
                        ' Literal/Wildcard line: split by commas outside quotes.
                        Dim tokens As New List(Of String)()
                        Dim sb As New StringBuilder()
                        Dim inQuotes As Boolean = False
                        For i As Integer = 0 To item.Length - 1
                            Dim ch As Char = item(i)
                            If ch = """"c Then
                                inQuotes = Not inQuotes
                                sb.Append(ch)
                            ElseIf ch = ","c AndAlso Not inQuotes Then
                                tokens.Add(sb.ToString().Trim())
                                sb.Clear()
                            Else
                                sb.Append(ch)
                            End If
                        Next
                        If sb.Length > 0 Then
                            tokens.Add(sb.ToString().Trim())
                        End If

                        For Each rawTok As String In tokens
                            Dim tok As String = rawTok.Trim()
                            If String.IsNullOrEmpty(tok) Then
                                Continue For
                            End If

                            Dim prefix As String = DEFAULT_PLACEHOLDER
                            Dim tokenWithoutMarker As String = tok
                            Dim markerStart As Integer = tok.IndexOf("{{")
                            Dim markerEnd As Integer = tok.IndexOf("}}")
                            If markerStart >= 0 AndAlso markerEnd > markerStart Then
                                Dim between As String = tok.Substring(markerStart + 2, markerEnd - markerStart - 2).Trim()
                                If Not String.IsNullOrEmpty(between) Then
                                    prefix = between
                                End If
                                tokenWithoutMarker = tok.Remove(markerStart, (markerEnd + 2) - markerStart).Trim()
                            End If

                            ' Build regex pattern: "*" → "[\p{L}\p{N}-]*"; quoted tokens and literals are escaped.
                            Dim patternText As String = String.Empty
                            If tokenWithoutMarker.StartsWith("""") AndAlso tokenWithoutMarker.EndsWith("""") AndAlso tokenWithoutMarker.Length >= 2 Then
                                Dim inner As String = tokenWithoutMarker.Substring(1, tokenWithoutMarker.Length - 2)
                                patternText = Regex.Escape(inner)
                            ElseIf tokenWithoutMarker.Contains("*") Then
                                Dim sbPat As New StringBuilder()
                                For Each c As Char In tokenWithoutMarker
                                    If c = "*"c Then
                                        sbPat.Append("[\p{L}\p{N}-]*")
                                    Else
                                        sbPat.Append(Regex.Escape(c.ToString()))
                                    End If
                                Next
                                patternText = sbPat.ToString()
                            Else
                                patternText = Regex.Escape(tokenWithoutMarker)
                            End If

                            Try
                                Dim rx As New Regex(patternText, RegexOptions.IgnoreCase Or RegexOptions.Compiled)
                                groupIDCounter += 1
                                patternInfos.Add(New PatternInfo(rx, prefix, groupIDCounter))
                            Catch rgEx As System.Exception
                                ShowCustomMessageBox($"Invalid pattern '{patternText}': {rgEx.Message}")
                            End Try
                        Next
                    End If
                Next

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error parsing anonymization file: {ex.Message}")
            End Try

            Return patternInfos
        End Function


        ' ------------------------------------------------------------------------
        ' Helper: BuildDefaultPromptFromFile(modelName) As String
        '   Returns a comma-separated list of literal/wildcard tokens (with {{prefix}} intact)
        '   from [All] and [ModelName], ignoring "Regex:" lines.
        ' ------------------------------------------------------------------------
        ''' <summary>
        ''' Builds a comma-separated default prompt string from file-based literal/wildcard entries
        ''' for <c>[All]</c> and matching model sections, ignoring <c>Regex:</c> lines.
        ''' </summary>
        ''' <param name="modelName">Model name used to determine whether a section applies.</param>
        ''' <returns>A comma-separated token list; empty string if file is missing or no applicable tokens exist.</returns>
        Private Function BuildDefaultPromptFromFile(ByVal modelName As String) As String
            Dim literals As New List(Of String)()

            Try
                If Not File.Exists(AnonFilepath) Then
                    Return String.Empty
                End If

                Dim lines As String() = File.ReadAllLines(AnonFilepath)
                Dim currentSection As String = String.Empty
                Dim isAllSection As Boolean = False
                Dim isModelSection As Boolean = False

                For Each rawLine As String In lines
                    Dim line As String = rawLine.Trim()
                    If line.StartsWith(";") OrElse String.IsNullOrEmpty(line) Then
                        Continue For
                    End If

                    If line.StartsWith("[") AndAlso line.EndsWith("]") Then
                        currentSection = line.Substring(1, line.Length - 2).Trim()
                        isAllSection = String.Equals(currentSection, "All", StringComparison.OrdinalIgnoreCase)

                        If Not isAllSection Then
                            Dim modelTokens As String() = currentSection.Split(","c)
                            Dim found As Boolean = False
                            For Each tok In modelTokens
                                If String.Equals(tok.Trim(), modelName, StringComparison.OrdinalIgnoreCase) Then
                                    found = True
                                    Exit For
                                End If
                            Next
                            isModelSection = found
                        Else
                            isModelSection = False
                        End If

                        Continue For
                    End If

                    If line.StartsWith("Anon", StringComparison.OrdinalIgnoreCase) Then
                        Continue For
                    End If

                    If (isAllSection OrElse isModelSection) AndAlso Not line.StartsWith("Regex:", StringComparison.OrdinalIgnoreCase) Then
                        ' Split by commas ignoring quoted segments:
                        Dim tokens As New List(Of String)()
                        Dim sb As New StringBuilder()
                        Dim inQuotes As Boolean = False
                        For i As Integer = 0 To line.Length - 1
                            Dim ch As Char = line(i)
                            If ch = """"c Then
                                inQuotes = Not inQuotes
                                sb.Append(ch)
                            ElseIf ch = ","c AndAlso Not inQuotes Then
                                tokens.Add(sb.ToString().Trim())
                                sb.Clear()
                            Else
                                sb.Append(ch)
                            End If
                        Next
                        If sb.Length > 0 Then
                            tokens.Add(sb.ToString().Trim())
                        End If

                        For Each tok As String In tokens
                            If Not String.IsNullOrWhiteSpace(tok) Then
                                literals.Add(tok.Trim())
                            End If
                        Next
                    End If
                Next

            Catch ex As System.Exception
                ShowCustomMessageBox($"Error building default prompt from file: {ex.Message}")
            End Try

            Return String.Join(", ", literals)
        End Function

    End Module

End Namespace