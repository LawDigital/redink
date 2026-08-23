' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.TextToSpeech.Commands.vb
' Purpose: Drives text-to-speech workflows for Word selections, imported documents, and PowerPoint speaker notes.
'
' Architecture:
'  - Entry Points: CreateAudio orchestrates user prompts, file ingestion, and downstream audio generation paths.
'  - File Import & Sanitization: Helper readers extract text from supported formats and insert sanitized content into a staging Word document.
'  - Voice & Output Configuration: Custom dialogs plus TTSSelectionForm gather voice pairs, languages, and output targets, persisting defaults via My.Settings.
'  - Conversation Normalization: NormalizeHostGuestConversation enforces deterministic H:/G: turn formatting for podcast-style transcripts.
'  - Speaker Notes Pipeline: GenerateAndPlayAudioFromSpeakerNotes cleans notes, requests audio, embeds MP3 media, and configures slide transitions.
'  - Progress & Cancellation: ProgressBarModule, ESC polling, and message boxes provide visibility and allow user-initiated aborts.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Diagnostics
Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Presentation
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Interop.Word
Imports NetOffice.PowerPointApi
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods
Imports System.Windows.Forms

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Creates an audio rendition from the current Word selection, imported files, JSON TTS definitions, or PowerPoint speaker notes.
    ''' </summary>
    Public Async Sub CreateAudio()
        If INILoadFail() Then Return

        DetectTTSEngines()

        ' — Nothing at all? bail out —
        If Not TTS_googleAvailable AndAlso Not TTS_openAIAvailable Then
            Return   ' no TTS provider configured
        End If

        Dim FilePath As String = ""
        Dim FromFile As String = ""
        SelectedText = ""

        Dim application As Word.Application = Globals.ThisAddIn.Application
        Dim selection As Microsoft.Office.Interop.Word.Selection = application.Selection
        If selection.Type = WdSelectionType.wdSelectionIP Then
            Dim answer As Integer = ShowCustomYesNoBox("You have not selected any text. Do you instead want to create audio from a document file or add audio to a powerpoint with speaker notes?", "Yes", "No")
            If answer <> 1 Then Return

            If INI_AllowLegacyDocFiles Then
                DragDropFormLabel = "Document files (.txt, .doc, .docx, .xlsx, .pdf), Powerpoint (.pptx), email (.msg, .eml)."
                DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Rich Text Files (*.rtf)|*.rtf|" &
                                 "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                 "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                 "PDF Files (*.pdf)|*.pdf|" &
                                 "PowerPoint Files (*.pptx)|*.pptx|" &
                                 "Email Files (*.msg;*.eml)|*.msg;*.eml"
            Else
                DragDropFormLabel = "Document files (.txt, .docx, .xlsx, .pdf), Powerpoint (.pptx), email (.msg, .eml)."
                DragDropFormFilter = "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Rich Text Files (*.rtf)|*.rtf|" &
                                 "Word Documents (*.docx)|*.docx|" &
                                 "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                 "PDF Files (*.pdf)|*.pdf|" &
                                 "PowerPoint Files (*.pptx)|*.pptx|" &
                                 "Email Files (*.msg;*.eml)|*.msg;*.eml"
            End If

            FilePath = GetFileName()
            DragDropFormLabel = ""
            DragDropFormFilter = ""
            If String.IsNullOrWhiteSpace(FilePath) Then
                ShowCustomMessageBox("No file has been selected - will abort.")
                Return
            End If

            Dim ext As String = IO.Path.GetExtension(FilePath).ToLowerInvariant()

            Select Case ext
                Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm"
                    FromFile = ReadTextFile(FilePath, True)
                Case ".rtf"
                    FromFile = ReadRtfAsText(FilePath, True)
                Case ".doc"
                    If INI_AllowLegacyDocFiles Then
                        FromFile = ReadWordDocument(FilePath, True)
                    Else
                        FromFile = "Error: .doc format disabled for security."
                    End If
                Case ".docx"
                    FromFile = ReadDocxSandboxed(FilePath)
                Case ".xlsx"
                    FromFile = ReadXlsxSandboxed(FilePath)
                Case ".pdf"
                    FromFile = Await ReadPdfAsText(FilePath, True, False, False, _context)
                Case ".pptx"
                    FromFile = "pptx"
                Case ".eml"
                    FromFile = ReadEmlSandboxed(FilePath)
                Case ".msg"
                    FromFile = ReadMsgSandboxed(FilePath)
                Case Else
                    FromFile = "Error: File type not supported."
            End Select

            If FromFile.StartsWith("Error:") Then
                ShowCustomMessageBox(FromFile)
                Return
            End If
            If String.IsNullOrWhiteSpace(FromFile) Then
                ShowCustomMessageBox("The file you provided did not contain any text - will abort.")
                Return
            End If
            If FromFile <> "pptx" Then

                Dim newDoc As Word.Document = Globals.ThisAddIn.Application.Documents.Add()
                newDoc.Activate()

                Dim rng As Word.Range = newDoc.Content
                rng.Collapse(Word.WdCollapseDirection.wdCollapseEnd)

                ' Sanitize: remove NULs and normalize line breaks to CRLF
                Dim safeText As String = If(FromFile, String.Empty)
                safeText = safeText.Replace(ChrW(0), String.Empty)

                rng.InsertAfter(safeText)

                newDoc.Content.Select()
                SelectedText = newDoc.Application.Selection.Text.Trim()

                answer = ShowCustomYesNoBox("The content of your document has been inserted into a new document. Continue with the audio generation?", "Yes", "No")
                If answer <> 1 Then Return

            End If
        Else
            SelectedText = selection.Text.Trim()
        End If
        If SelectedText.Contains("H: ") And SelectedText.Contains("G: ") Then
            ReadPodcast(SelectedText)
        Else
            If selection.Text.Trim().StartsWith("{") Then
                Dim selectedoutputpath As String = (If(String.IsNullOrEmpty(My.Settings.TTSOutputPath), System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile), My.Settings.TTSOutputPath))
                selectedoutputpath = ShowCustomInputBox("Where should the audio generated from your JSON TTS file be saved to?", $"{AN} Create Audiobook", True, selectedoutputpath)
                If String.IsNullOrWhiteSpace(selectedoutputpath) Then
                    ' Use default path (Desktop) with default filename
                    selectedoutputpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), TTSDefaultFile)
                ElseIf selectedoutputpath.EndsWith("\") OrElse selectedoutputpath.EndsWith("/") Then
                    ' If only a folder is given, append default filename
                    selectedoutputpath = System.IO.Path.Combine(selectedoutputpath, TTSDefaultFile)
                Else
                    Dim dir As String = System.IO.Path.GetDirectoryName(selectedoutputpath)
                    Dim fileName As String = System.IO.Path.GetFileName(selectedoutputpath)

                    ' If no directory is found, assume Desktop as the base
                    If String.IsNullOrWhiteSpace(dir) Then
                        selectedoutputpath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), fileName)
                        dir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                    End If

                    ' If no filename is given, use the default filename
                    If String.IsNullOrWhiteSpace(fileName) Then
                        selectedoutputpath = System.IO.Path.Combine(dir, TTSDefaultFile)
                    End If

                    ' Ensure the filename has ".mp3" extension
                    If Not fileName.EndsWith(".mp3", StringComparison.OrdinalIgnoreCase) Then
                        selectedoutputpath = System.IO.Path.Combine(dir, fileName & ".mp3")
                    End If
                End If
                GenerateAndPlayAudio(selection.Text, selectedoutputpath, "", "")
                Return
            Else
                Dim Voices As Integer = ShowCustomYesNoBox("Do you want to use alternate voices to read the text?", "No, one voice", "Yes, alternate", "Create Audio")
                If Voices = 0 Then Return
                Using frm As New TTSSelectionForm("Select the voice you wish To use For creating your audio file And configure where To save it.", $"{AN} Text-To-Speech - Select Voices", Voices = 2, FromFile = "pptx") ' TTSSelectionForm(_context, INI_OAuth2ClientMail, INI_OAuth2Scopes, INI_APIKey, INI_OAuth2Endpoint, INI_OAuth2ATExpiry, "Select the voice you wish To use For creating your audio file And configure where To save it.", $"{AN} Text-To-Speech - Select Voices", Voices = 2)
                    Dim __safeDialogOwner187 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
                    If If(__safeDialogOwner187 IsNot Nothing, frm.ShowDialog(__safeDialogOwner187), frm.ShowDialog()) = DialogResult.OK Then
                        Dim selectedVoices As List(Of String) = frm.SelectedVoices
                        Dim selectedLanguage As String = frm.SelectedLanguage
                        Dim outputPath As String = frm.SelectedOutputPath
                        If FromFile = "pptx" Then
                            GenerateAndPlayAudioFromSpeakerNotes(FilePath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), If(Voices = 2, selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), ""))
                        Else
                            GenerateAndPlayAudioFromSelectionParagraphs(outputPath, selectedLanguage, selectedVoices(0).Replace(" (male)", "").Replace(" (female)", ""), If(Voices = 2, selectedVoices(1).Replace(" (male)", "").Replace(" (female)", ""), ""))
                        End If
                    End If
                End Using
            End If
        End If
    End Sub

    ''' <summary>
    ''' Normalizes an LLM result into alternating host (H) and guest (G) turns with clean spacing and blank-line separation.
    ''' </summary>
    ''' <param name="llmResult">Raw conversation text returned by the model.</param>
    ''' <returns>Formatted conversation containing only populated turns with Windows line endings.</returns>
    Public Function NormalizeHostGuestConversation(llmResult As String) As String
        ' Ensures:
        ' 1) Conversation consists of turns starting with "H:" or "G:".
        ' 2) Any newline directly after "H:" / "G:" is removed.
        ' 3) A space follows the speaker tag: "H: " / "G: ".
        ' 4) Turns are separated by exactly one blank line (i.e. two consecutive CRLFs).
        ' 5) Trailing blank lines in each turn's body are trimmed so we never end up with more than one empty line between turns.
        If String.IsNullOrWhiteSpace(llmResult) Then Return String.Empty

        ' Normalize line endings to LF for regex processing
        Dim s = llmResult.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        ' Collapse 3+ blank lines to 2 (optional hygiene)
        s = System.Text.RegularExpressions.Regex.Replace(s, "\n{3,}", vbLf & vbLf)

        ' Regex: capture each turn starting with H: or G:
        Dim rx As New System.Text.RegularExpressions.Regex("(?<=^|\n)(H:|G:)(?:\s*\n)?(.*?)(?=\n(?:H:|G:)|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim matches = rx.Matches(s)

        Dim turns As New List(Of String)

        For Each m As System.Text.RegularExpressions.Match In matches
            Dim tag = m.Groups(1).Value ' "H:" or "G:"
            Dim body = m.Groups(2).Value

            ' Trim leading whitespace/newlines at start of body
            body = System.Text.RegularExpressions.Regex.Replace(body, "^\s+", "")

            ' Trim trailing whitespace/newlines to avoid extra blank lines between turns
            body = System.Text.RegularExpressions.Regex.Replace(body, "\s+$", "")

            ' Skip empty bodies (optional)
            If String.IsNullOrWhiteSpace(body) Then Continue For

            ' Ensure tag followed by single space
            Dim normalizedTag = tag & " "

            ' Restore Windows line endings inside body
            body = body.Replace(vbLf, vbCrLf)

            ' Remove any accidental CR/LF immediately after tag (defensive)
            If body.StartsWith(vbCrLf) Then body = body.TrimStart()

            turns.Add(normalizedTag & body)
        Next

        ' Join turns with exactly one blank line (two CRLFs)
        Return String.Join(vbCrLf & vbCrLf, turns)
    End Function

    ''' <summary>
    ''' Generates audio for each slide that contains speaker notes and embeds the resulting MP3 into the presentation.
    ''' </summary>
    ''' <param name="presentationFilePath">Full path to the PowerPoint file that will be updated.</param>
    ''' <param name="languageCode">Language/locale identifier passed to the TTS provider.</param>
    ''' <param name="voiceName">Primary voice used for the first applicable segment.</param>
    ''' <param name="voiceNameAlt">Optional alternate voice, toggled on successive slides when supplied.</param>
    Public Async Function GenerateAndPlayAudioFromSpeakerNotes(
        presentationFilePath As String,
        Optional languageCode As String = "en-US",
        Optional voiceName As String = "en-US-Studio-O",
        Optional voiceNameAlt As String = ""
    ) As System.Threading.Tasks.Task

        Dim ppApp As NetOffice.PowerPointApi.Application = Nothing
        Dim presentation As NetOffice.PowerPointApi.Presentation = Nothing

        Try
            '--- Load and save TTS settings ---
            Dim NoSSML As Boolean = My.Settings.NoSSML
            Dim Pitch As Double = My.Settings.Pitch
            Dim SpeakingRate As Double = My.Settings.Speakingrate
            Dim CleanText As Boolean = False
            Dim CleanTextPrompt As String = My.Settings.CleanTextPrompt
            If String.IsNullOrWhiteSpace(CleanTextPrompt) Then CleanTextPrompt = SP_CleanTextPrompt

            Dim params() As SLib.InputParameter = {
                New SLib.InputParameter("Pitch", Pitch),
                New SLib.InputParameter("Speaking Rate", SpeakingRate),
                New SLib.InputParameter("No SSML", NoSSML),
                New SLib.InputParameter("Clean text", CleanText)
            }
            If Not ShowCustomVariableInputForm("Parameters for audio generation:", "Create Audio (Slides)", params) Then
                Return
            End If

            If CleanText Then
                CleanTextPrompt = ShowCustomInputBox("Please enter the prompt to 'clean' the text with (each paragraph will be submitted to this prompt)", "Create Audio", False, CleanTextPrompt, Context:=_context).Trim()
                If CleanTextPrompt = "ESC" Then Return
                If CleanTextPrompt = "" Then
                    CleanText = False
                Else
                    My.Settings.CleanTextPrompt = CleanTextPrompt
                    My.Settings.Save()
                End If
            End If

            Pitch = CDbl(params(0).Value)
            SpeakingRate = CDbl(params(1).Value)
            NoSSML = CBool(params(2).Value)
            CleanText = CBool(params(3).Value)
            My.Settings.NoSSML = NoSSML
            My.Settings.Pitch = Pitch
            My.Settings.Speakingrate = SpeakingRate
            My.Settings.Save()

            Dim useAlternate As Boolean = (voiceNameAlt <> "")
            Dim currentVoice As String = voiceName
            Dim firstUsed As Boolean = False

            '--- Open PowerPoint ---
            ppApp = New NetOffice.PowerPointApi.Application()
            presentation = ppApp.Presentations.Open(
                presentationFilePath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse)

            If presentation.Slides.Count > 0 Then

                ShowProgressBarInSeparateThread($"{AN} Audio Generation", "Starting audio generation...")
                ProgressBarModule.CancelOperation = False
                GlobalProgressValue = 0
                GlobalProgressMax = presentation.Slides.Count

                For slideIndex As Integer = 1 To presentation.Slides.Count

                    If (GetAsyncKeyState(VK_ESCAPE) And &H8000) <> 0 Or (GetAsyncKeyState(VK_ESCAPE) And 1) <> 0 Or ProgressBarModule.CancelOperation Then
                        ShowCustomMessageBox("Audio generation aborted by user.")
                        ProgressBarModule.CancelOperation = True
                        presentation.Close()
                        ppApp.Quit()
                        Return
                    End If

                    Dim slide As NetOffice.PowerPointApi.Slide = presentation.Slides(slideIndex)

                    ' 1) Find the notes-placeholder
                    Dim notesShape As NetOffice.PowerPointApi.Shape = Nothing
                    For i As Integer = 1 To slide.NotesPage.Shapes.Placeholders.Count
                        Dim shp = slide.NotesPage.Shapes.Placeholders(i)
                        If shp.PlaceholderFormat.Type = PpPlaceholderType.ppPlaceholderBody Then
                            notesShape = shp
                            Exit For
                        End If
                    Next
                    If notesShape Is Nothing Then Continue For

                    Dim notesText As String = notesShape.TextFrame.TextRange.Text.Trim()
                    If String.IsNullOrWhiteSpace(notesText) Then Continue For
                    If Not notesText.EndsWith(".") Then notesText &= "."

                    ' switch voice if needed
                    If useAlternate Then
                        If Not firstUsed Then
                            firstUsed = True
                        Else
                            currentVoice = If(currentVoice = voiceName, voiceNameAlt, voiceName)
                        End If
                    End If

                    If CleanText Then
                        ' Remove any unwanted characters from the paragraph text.
                        notesText = Await LLM(CleanTextPrompt, "<TEXTTOPROCESS>" & notesText & "</TEXTTOPROCESS>", "", "", 0, False, True)
                        notesText = notesText.Trim().Replace("<TEXTTOPROCESS>", "").Replace("</TEXTTOPROCESS>", "").Trim()
                        Debug.WriteLine("Cleaned notes = " & notesText & vbCrLf & vbCrLf)

                    End If

                    '--- Get audio bytes from TTS ---
                    Dim audioBytes As Byte() = Await GenerateAudioFromText(
                        notesText,
                        languageCode,
                        currentVoice,
                        NoSSML,
                        Pitch,
                        SpeakingRate,
                        "Slide " & slideIndex.ToString()
                    )
                    If audioBytes Is Nothing OrElse audioBytes.Length = 0 Then
                        Debug.WriteLine("[Debug] Slide " & slideIndex & ": no audio returned.")
                        Continue For
                    End If

                    '--- Save raw bytes as MP3 ---
                    Dim tempFile As String = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        $"ppt_audio_slide_{slideIndex}.mp3"
                    )
                    File.WriteAllBytes(tempFile, audioBytes)

                    Dim audioDurationSeconds As Double
                    Using mp3Reader As New NAudio.Wave.Mp3FileReader(tempFile)
                        audioDurationSeconds = mp3Reader.TotalTime.TotalSeconds
                    End Using

                    ' debug info
                    Dim exists As Boolean = File.Exists(tempFile)
                    Dim size As Long = If(exists, (New FileInfo(tempFile)).Length, -1L)
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: tempFile='{tempFile}', Exists={exists}, Size={size}")

                    Dim beforeCount = slide.Shapes.Count
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: Shapes before insert = {beforeCount}")

                    '--- Insert the MP3 ---
                    Dim mediaShape As NetOffice.PowerPointApi.Shape = Nothing
                    Try
                        mediaShape = slide.Shapes.AddMediaObject2(
                            fileName:=tempFile,
                            linkToFile:=MsoTriState.msoFalse,
                            saveWithDocument:=MsoTriState.msoTrue,
                            left:=10, top:=10,
                            width:=10, height:=10
                        )
                        Debug.WriteLine($"[Debug] AddMediaObject2 succeeded: Id={mediaShape.Id}, Type={mediaShape.Type}")

                        If mediaShape IsNot Nothing Then

                            '--- Determine the audio length in seconds ---   ' Playback on entry + hide while not playing
                            With mediaShape.AnimationSettings.PlaySettings
                                .PlayOnEntry = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                                .HideWhileNotPlaying = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                            End With

                            '--- Add a 1-second delay before audio playback ---
                            With mediaShape.AnimationSettings
                                .AdvanceMode = NetOffice.PowerPointApi.Enums.PpAdvanceMode.ppAdvanceOnTime
                                .AdvanceTime = 1     ' Seconds until playback
                            End With

                            '--- Auto-advance slide after (1s delay + audio length + 1s hold) ---
                            With slide.SlideShowTransition
                                .AdvanceOnTime = NetOffice.OfficeApi.Enums.MsoTriState.msoTrue
                                .AdvanceTime = CSng(audioDurationSeconds + 1.0)   ' 1s after end
                                .AdvanceOnClick = NetOffice.OfficeApi.Enums.MsoTriState.msoFalse
                            End With

                        End If

                    Catch comEx As System.Runtime.InteropServices.COMException
                        Debug.WriteLine($"[Error] COMException in AddMediaObject2: {comEx}")
                        Continue For
                    End Try

                    '--- Configure play settings and initially hide upon play ---
                    With mediaShape.AnimationSettings.PlaySettings
                        .PlayOnEntry = MsoTriState.msoTrue
                        .HideWhileNotPlaying = MsoTriState.msoTrue
                    End With

                    Dim afterCount = slide.Shapes.Count
                    Debug.WriteLine($"[Debug] Slide {slideIndex}: Shapes after insert = {afterCount}")

                    ' Update the current progress value and status label.
                    GlobalProgressValue = slideIndex
                    GlobalProgressLabel = $"Slide {slideIndex} of {GlobalProgressMax} (some may be skipped)"

                    If File.Exists(tempFile) Then
                        Try
                            File.Delete(tempFile)
                        Catch ex As System.Exception
                            Debug.WriteLine($"[Warning] Could not delete temp file '{tempFile}': {ex.Message}")
                        End Try
                    End If

                Next

                ' save & clean up

                ProgressBarModule.CancelOperation = True

                Try
                    presentation.Save()
                    Debug.WriteLine("[Debug] Presentation.Save succeeded.")
                Catch comSaveEx As System.Runtime.InteropServices.COMException
                    Debug.WriteLine($"[Debug] Presentation.Save failed: {comSaveEx.Message}")
                    ' PowerPoint reported write protection → overwrite via SaveAs
                    presentation.SaveAs(
                    presentationFilePath,
                    NetOffice.PowerPointApi.Enums.PpSaveAsFileType.ppSaveAsOpenXMLPresentation)
                    Debug.WriteLine("[Debug] Presentation.SaveAs succeeded.")
                End Try

                presentation.Close()
                ppApp.Quit()

                ShowCustomMessageBox("All slides with speaker notes have been amended with audio and auto-play.")

            Else
                ShowCustomMessageBox("No slides found in the presentation.")
                Return
            End If

        Catch ex As Exception
            Debug.WriteLine("[Error] Unexpected error: " & ex.ToString())
            ShowCustomMessageBox($"An unexpected error occurred when adding audio to the the slides ({ex.GetType().Name}): {ex.Message}")
        Finally
            If presentation IsNot Nothing Then presentation.Dispose()
            If ppApp IsNot Nothing Then ppApp.Dispose()
        End Try
    End Function

End Class
