' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FileRenamer.vb
' Purpose: Renames supported files in a user-selected folder using an LLM-generated
'          filename based on file content and optional reference document.
'
' Architecture:
'  - Instruction Library: Loads local/global *.txt instruction sets; each line "Title|Instruction"
'    (lines starting with ";" ignored). Manual instruction overrides prepared instruction.
'  - UI Parameter Collection: Uses SharedMethods.InputParameter and ShowCustomVariableInputForm
'    to collect rename parameters (prepared/manual instruction, reference doc usage, OCR, output language,
'    optional secondary model).
'  - Model Switching: Optional alternate model or secondary API selection based on configuration flags.
'  - Reference Document (optional): User selects a file; content loaded (with optional OCR) and tagged
'    as <REFERENCEDOC> for prompt enrichment.
'  - File Enumeration: Filters extensions (.pdf, .docx, .doc, .txt, .rtf, .ini, .csv, .log, .json, .xml, .html, .ht).
'  - Prompt Construction: System prompt via InterpolateAtRuntime(SP_Rename); user prompt encloses truncated
'    file content (<TEXTTOPROCESS>) and optional reference (<REFERENCEDOC>).
'  - LLM Invocation: Calls LLM(systemPrompt, userPrompt, ...) expecting a filename (body only or full).
'  - Sanitization & Uniqueness: Removes invalid characters (SanitizeFileName) and ensures uniqueness
'    (BuildUniqueTargetPath) with numeric suffixes.
'  - Progress & Cancellation: ProgressBarModule tracks state; user can abort mid-iteration.
'  - Summary: Provides mapping and error report; results copied to clipboard when applicable.
'  - External Dependencies: SharedLibrary.SharedMethods supplies UI, interpolation, LLM, progress,
'    clipboard, text editor, and model selection helpers.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Excel
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Async entry point that renames supported files in a user-selected folder using LLM output.
    ''' Steps: load instruction sets; collect parameters; optional reference document; enumerate files;
    ''' build system/user prompts; invoke LLM; sanitize and apply rename; show summary.
    ''' </summary>
    ''' <remarks>
    ''' Uses SharedMethods for UI dialogs, progress handling, prompt interpolation, LLM invocation,
    ''' clipboard operations, and model selection. Supports optional OCR and alternate model/API.
    ''' </remarks>
    Public Async Sub RenameDocumentsWithAi()
        Dim useSecondApi As Boolean = False
        Dim do2ndModel As Boolean = False

        Try
            ' ------------------------------
            ' 1. Load instruction library (prepared sets)
            ' ------------------------------
            Dim displayToInstruction As New System.Collections.Generic.Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
            Dim displayOptions As New System.Collections.Generic.List(Of String)()

            Dim localPath = ExpandEnvironmentVariables(INI_RenameLibPathLocal)
            Dim globalPath = ExpandEnvironmentVariables(INI_RenameLibPath)

            Dim EnumerateInstructionFiles As Func(Of String, System.Collections.Generic.IEnumerable(Of String)) =
                Function(p As String)
                    Dim result As New System.Collections.Generic.List(Of String)
                    If String.IsNullOrWhiteSpace(p) Then Return result
                    Try
                        If Directory.Exists(p) Then
                            result.AddRange(Directory.GetFiles(p, "*.txt", SearchOption.TopDirectoryOnly))
                        ElseIf File.Exists(p) Then
                            result.Add(p)
                        End If
                    Catch
                    End Try
                    Return result
                End Function

            Dim LoadFromFile As Action(Of String, Boolean) =
                Sub(file As String, isLocal As Boolean)
                    Try
                        For Each rawLine In System.IO.File.ReadAllLines(file)
                            Dim line = If(rawLine, "").Trim()
                            If line.Length = 0 OrElse line.StartsWith(";", StringComparison.Ordinal) Then Continue For
                            Dim parts = line.Split("|"c)
                            If parts.Length < 1 Then Continue For
                            Dim title = parts(0).Trim()
                            Dim instr As String = If(parts.Length >= 2, parts(1).Trim(), "")
                            If title.Length = 0 Then Continue For
                            Dim display = title & If(isLocal, " (local)", "")
                            Dim unique = MakeUniqueDisplay(display, displayToInstruction.Keys)
                            displayToInstruction(unique) = instr
                            displayOptions.Add(unique)
                        Next
                    Catch
                    End Try
                End Sub

            For Each f In EnumerateInstructionFiles(localPath) : LoadFromFile(f, True) : Next
            For Each f In EnumerateInstructionFiles(globalPath) : LoadFromFile(f, False) : Next

            ' ------------------------------
            ' 2. Retrieve defaults from settings
            ' ------------------------------
            Dim defaultManual As String = ""
            Dim defaultOutputLanguage As String = ""
            Try : defaultManual = System.Convert.ToString(My.Settings.Rename_ManualInstruction) : Catch : End Try
            Try : defaultOutputLanguage = System.Convert.ToString(My.Settings.Rename_OutputLanguage) : Catch : End Try

            Dim manualInstruction = defaultManual
            Dim UserOutputLanguage = defaultOutputLanguage
            Dim useReferenceDoc As Boolean = False
            Dim doOcr As Boolean = False
            Dim referenceDocPath As String = Nothing

            Dim hasSecondary = (Not String.IsNullOrWhiteSpace(INI_AlternateModelPath)) OrElse INI_SecondAPI

            ' ------------------------------
            ' 3. Build parameter objects (UI)
            ' ------------------------------
            Dim defaultPreparedDisplay = If(displayOptions.Count > 0, displayOptions(0), "<<disabled>>")
            Dim p0 As New SLib.InputParameter("Prepared instruction set", defaultPreparedDisplay)
            If displayOptions.Count > 0 Then p0.Options = New System.Collections.Generic.List(Of String)(displayOptions)

            Dim p1 As SLib.InputParameter = If(displayOptions.Count = 0,
                                               New SLib.InputParameter("Manual instruction (required if no library)", manualInstruction),
                                               New SLib.InputParameter("Manual instruction (overrides)", manualInstruction))
            Dim pRef As New SLib.InputParameter("Also use reference document", useReferenceDoc)
            Dim pOcr As New SLib.InputParameter("Do OCR if needed (PDFs)", doOcr)
            Dim pLang As New SLib.InputParameter("Output language", UserOutputLanguage)
            Dim pAlt As SLib.InputParameter = Nothing
            If hasSecondary Then
                do2ndModel = False
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    pAlt = New SLib.InputParameter("Use a secondary model", do2ndModel)
                Else
                    pAlt = New SLib.InputParameter("Use the secondary model", do2ndModel)
                End If
            End If

            Dim params() As SLib.InputParameter =
                If(hasSecondary,
                   New SLib.InputParameter() {p0, p1, pRef, pOcr, pLang, pAlt},
                   New SLib.InputParameter() {p0, p1, pRef, pOcr, pLang})

            Dim extraText As String = Nothing
            Dim extraAction As System.Action = Nothing
            Dim closeAfterExtra As Boolean = False

            If Not String.IsNullOrWhiteSpace(localPath) Then
                extraText = "Edit Local Library"
                extraAction =
                        Sub()
                            Try
                                ' Create file with sample content if it doesn't exist or contains only whitespace
                                Dim needsSampleContent As Boolean = False
                                If Not File.Exists(localPath) Then
                                    needsSampleContent = True
                                Else
                                    Try
                                        Dim content As String = File.ReadAllText(localPath, System.Text.Encoding.UTF8)
                                        needsSampleContent = String.IsNullOrWhiteSpace(content)
                                    Catch
                                        needsSampleContent = True
                                    End Try
                                End If

                                If needsSampleContent Then
                                    Try
                                        File.WriteAllText(localPath,
                                        "; Red Ink File Renamer - Local Library" & vbCrLf &
                                        "; Format: Title | Instruction" & vbCrLf &
                                        "; Lines starting with ; are comments" & vbCrLf &
                                        vbCrLf &
                                        "Contract Naming|Rename to format: [Date YYYY-MM-DD] [PartyA] - [PartyB] - [Contract Type]. Extract the effective date, main contracting parties, and contract type (e.g., NDA, Service Agreement, License)." & vbCrLf & vbCrLf &
                                        "Invoice Naming|Rename to format: [Date YYYY-MM-DD] [Vendor] - Invoice [Number] - [Amount]. Extract invoice date, vendor name, invoice number, and total amount." & vbCrLf & vbCrLf &
                                        "Correspondence|Rename to format: [Date YYYY-MM-DD] [Sender] to [Recipient] - [Subject]. Extract the letter date, sender, recipient, and brief subject." & vbCrLf & vbCrLf &
                                        "Legal Filing|Rename to format: [Case Number] - [Document Type] - [Filing Date YYYY-MM-DD]. Extract case reference, document type (Motion, Brief, Order), and filing date." & vbCrLf & vbCrLf &
                                        "Meeting Minutes|Rename to format: [Date YYYY-MM-DD] [Meeting Type] Minutes - [Key Topic]. Extract meeting date, type (Board, Team, Client), and primary topic discussed." & vbCrLf & vbCrLf &
                                        "Report Naming|Rename to format: [Date YYYY-MM-DD] [Report Type] - [Subject] - [Author]. Extract report date, type (Financial, Technical, Status), subject, and author if available." & vbCrLf,
                                        System.Text.Encoding.UTF8)
                                    Catch ex As Exception
                                        SLib.ShowCustomMessageBox($"Cannot create file: {ex.Message}")
                                        Return
                                    End Try
                                End If
                                ' Open the local library in the editor
                                SLib.ShowTextFileEditor(localPath, $"{AN} Local Library '{localPath}':", False, _context)
                            Catch ex As Exception
                                SLib.ShowCustomMessageBox("Error while opening the local library:" & vbCrLf & ex.Message)
                                Exit Sub
                            End Try

                            ' Inform the user about activation timing
                            SLib.ShowCustomMessageBox("Any changes to the local library will only be active the next time this feature is called up.")
                        End Sub
            End If

            If ShowCustomVariableInputForm("Please set the rename parameters:", AN & " AI File Renamer", params, extraButtonText:=extraText,
                                                                                        extraButtonAction:=extraAction,
                                                                                        CloseAfterExtra:=closeAfterExtra) = False Then Return

            Dim chosenPreparedDisplay = System.Convert.ToString(params(0).Value)
            manualInstruction = System.Convert.ToString(params(1).Value)
            Try : useReferenceDoc = System.Convert.ToBoolean(params(2).Value) : Catch : useReferenceDoc = False : End Try
            Try : doOcr = System.Convert.ToBoolean(params(3).Value) : Catch : doOcr = False : End Try
            UserOutputLanguage = System.Convert.ToString(params(4).Value)
            If hasSecondary Then
                Try : do2ndModel = System.Convert.ToBoolean(params(5).Value) : Catch : do2ndModel = False : End Try
            End If

            ' ------------------------------
            ' 4. Persist settings
            ' ------------------------------
            Try : My.Settings.Rename_ManualInstruction = manualInstruction : Catch : End Try
            Try : My.Settings.Rename_OutputLanguage = UserOutputLanguage : Catch : End Try
            Try : My.Settings.Save() : Catch : End Try

            ' Determine effective instruction (manual overrides library selection)
            Dim preparedInstruction As String = Nothing
            If chosenPreparedDisplay IsNot Nothing Then displayToInstruction.TryGetValue(chosenPreparedDisplay, preparedInstruction)
            Dim manualOverrides = Not String.IsNullOrWhiteSpace(manualInstruction)
            Dim preparedMissingInstruction = (Not manualOverrides) AndAlso Not String.IsNullOrWhiteSpace(chosenPreparedDisplay) AndAlso String.IsNullOrWhiteSpace(preparedInstruction)

            If manualOverrides Then
                preparedInstruction = Nothing
                chosenPreparedDisplay = Nothing
            End If

            Dim effectiveInstruction = If(manualOverrides,
                                          manualInstruction,
                                          If(preparedMissingInstruction,
                                             "Propose a concise, meaningful filename (English) based on document content.",
                                             preparedInstruction))
            If String.IsNullOrWhiteSpace(effectiveInstruction) Then
                ShowCustomMessageBox("No rename instruction provided.")
                Return
            End If

            ' ------------------------------
            ' 5. Optional model switching
            ' ------------------------------
            If hasSecondary AndAlso do2ndModel Then
                If Not String.IsNullOrWhiteSpace(INI_AlternateModelPath) Then
                    If Not ShowModelSelection(_context, INI_AlternateModelPath) Then
                        originalConfigLoaded = False
                        ShowCustomMessageBox("The alternate model could not be loaded - aborting.")
                        Return
                    Else
                        useSecondApi = True
                    End If
                ElseIf INI_SecondAPI Then
                    useSecondApi = True
                End If
            End If

            ' ------------------------------
            ' 6. If reference doc requested, prompt for file
            ' ------------------------------
            Dim referenceContent As String = ""
            If useReferenceDoc Then
                Try
                    DragDropFormLabel = ""
                    DragDropFormFilter = ""
                    referenceContent = Await GetFileContent(Nothing, False, Not String.IsNullOrWhiteSpace(INI_APICall_Object), True)
                    If String.IsNullOrWhiteSpace(referenceContent) Then
                        ShowCustomMessageBox("The file you have selected is empty or not supported - exiting.")
                        Return
                    End If
                Catch ex As Exception
                    ShowCustomMessageBox("Reference document selection failed: " & ex.Message)
                    Return
                End Try
            End If

            ' ------------------------------
            ' 7. User selects directory
            ' ------------------------------
            Dim selectedFolder As String = Nothing
            Try
                Using dlg As New FolderBrowserDialog()
                    dlg.Description = "Select folder with files to rename"
                    dlg.ShowNewFolderButton = False
                    If dlg.ShowDialog() <> DialogResult.OK OrElse String.IsNullOrWhiteSpace(dlg.SelectedPath) Then
                        ShowCustomMessageBox("No folder selected.")
                        Return
                    End If
                    selectedFolder = dlg.SelectedPath
                End Using
            Catch ex As Exception
                ShowCustomMessageBox("Folder selection failed: " & ex.Message)
                Return
            End Try

            Dim files() As String = {}
            Try
                files = Directory.GetFiles(selectedFolder, "*.*", SearchOption.TopDirectoryOnly).
                Where(Function(f) {".pdf", ".docx", ".doc", ".txt", ".rtf", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".ht"}.Contains(Path.GetExtension(f).ToLowerInvariant())).ToArray()
            Catch ex As Exception
                ShowCustomMessageBox("Failed to enumerate files: " & ex.Message)
                Return
            End Try
            If files.Length = 0 Then
                ShowCustomMessageBox("Folder contains no supported files.")
                Return
            End If

            ' ------------------------------
            ' 8. Assign globals (prompt & language)
            ' ------------------------------
            Globals.ThisAddIn.OtherPrompt = effectiveInstruction
            Globals.ThisAddIn.OutputLanguage = UserOutputLanguage

            ' ------------------------------
            ' 9. Progress bar initialization
            ' ------------------------------
            ShowProgressBarInSeparateThread(AN & " AI Rename", "Renaming files...")
            ProgressBarModule.CancelOperation = False
            GlobalProgressMax = files.Length
            GlobalProgressValue = 0
            GlobalProgressLabel = "Starting..."

            ' ------------------------------
            ' 10. Iterate and rename
            ' ------------------------------
            Dim renamedCount As Integer = 0
            Dim errorList As New System.Collections.Generic.List(Of String)
            Dim mapping As New System.Collections.Generic.List(Of String)

            For i = 0 To files.Length - 1
                If ProgressBarModule.CancelOperation Then Exit For

                Dim filePath = files(i)
                FileNameBody = Path.GetFileNameWithoutExtension(filePath)
                Dim extension = Path.GetExtension(filePath)

                Try
                    Dim creationDate = File.GetCreationTime(filePath)
                    Dim modificationDate = File.GetLastWriteTime(filePath)
                    FileDate = "FileCreationDate=" & creationDate.ToString("yyyy-MM-dd HH:mm:ss") &
                               " FileModificationDate=" & modificationDate.ToString("yyyy-MM-dd HH:mm:ss")
                Catch
                    FileDate = "[not provided]"
                End Try

                GlobalProgressValue = i + 1
                GlobalProgressLabel = "Reading: " & Path.GetFileName(filePath)

                Dim fileContent As String = ""
                Try
                    fileContent = Await GetFileContent(filePath, True, doOcr, False)
                Catch ex As Exception
                    errorList.Add(Path.GetFileName(filePath) & " -> read failed: " & ex.Message)
                    Continue For
                End Try
                If String.IsNullOrWhiteSpace(fileContent) Then
                    errorList.Add(Path.GetFileName(filePath) & " -> empty or unreadable content")
                    Continue For
                End If

                ' Build user prompt with tags
                Dim truncated = TruncateForPrompt(fileContent, 60000)
                Dim refTruncated = If(useReferenceDoc AndAlso Not String.IsNullOrWhiteSpace(referenceContent),
                                      TruncateForPrompt(referenceContent, 40000),
                                      "")

                Dim userPromptBuilder As New StringBuilder()
                userPromptBuilder.AppendLine("<TEXTTOPROCESS>")
                userPromptBuilder.AppendLine(truncated)
                userPromptBuilder.AppendLine("</TEXTTOPROCESS>")
                If useReferenceDoc Then
                    userPromptBuilder.AppendLine("<REFERENCEDOC>")
                    userPromptBuilder.AppendLine(refTruncated)
                    userPromptBuilder.AppendLine("</REFERENCEDOC>")
                End If

                Dim userPrompt = userPromptBuilder.ToString()

                ' System prompt via interpolation (includes OtherPrompt, OutputLanguage, FileNameBody automatically)
                Dim systemPrompt As String = ""
                Try
                    systemPrompt = InterpolateAtRuntime(SP_Rename)
                Catch ex As Exception
                    errorList.Add(Path.GetFileName(filePath) & " -> system prompt build failed: " & ex.Message)
                    Continue For
                End Try
                If String.IsNullOrWhiteSpace(systemPrompt) Then
                    errorList.Add(Path.GetFileName(filePath) & " -> empty system prompt")
                    Continue For
                End If

                GlobalProgressLabel = "Renaming: " & Path.GetFileName(filePath)

                Dim llmFilenameRaw As String = ""
                Try
                    ' Expect the LLM to return ONLY the new filename body (without extension) or full name.
                    llmFilenameRaw = Await LLM(systemPrompt, userPrompt, "", "", 0, useSecondApi, False)
                    llmFilenameRaw = llmFilenameRaw.Replace(vbCr, " ").Replace(vbLf, " ").Trim()
                Catch ex As Exception
                    errorList.Add(Path.GetFileName(filePath) & " -> LLM error: " & ex.Message)
                    Continue For
                End Try

                If String.IsNullOrWhiteSpace(llmFilenameRaw) Then
                    errorList.Add(Path.GetFileName(filePath) & " -> LLM returned empty name")
                    Continue For
                End If

                Dim sanitized = SanitizeFileName(llmFilenameRaw.Trim())
                If sanitized.EndsWith(extension, StringComparison.OrdinalIgnoreCase) Then
                    sanitized = Path.GetFileNameWithoutExtension(sanitized)
                End If
                If String.IsNullOrWhiteSpace(sanitized) Then
                    errorList.Add(Path.GetFileName(filePath) & " -> sanitized filename empty")
                    Continue For
                End If

                Dim targetPath = BuildUniqueTargetPath(selectedFolder, sanitized, extension)
                Try
                    File.Move(filePath, targetPath)
                    mapping.Add(Path.GetFileName(filePath) & " => " & Path.GetFileName(targetPath))
                    renamedCount += 1
                Catch ex As Exception
                    errorList.Add(Path.GetFileName(filePath) & " -> rename failed: " & ex.Message)
                End Try
            Next

            ProgressBarModule.CancelOperation = True

            ' ------------------------------
            ' 11. Summary report
            ' ------------------------------
            Dim sb As New StringBuilder()
            Dim sb2 As New StringBuilder()
            sb.AppendLine("Processed files: " & files.Length)
            sb.AppendLine("Renamed successfully: " & renamedCount)
            Dim aborted = (GlobalProgressValue < files.Length)
            If aborted Then sb.AppendLine("Operation aborted by user after " & GlobalProgressValue & " files.")
            If mapping.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("Mappings:")
                sb.AppendLine(String.Join(vbCrLf, mapping))
                sb2.AppendLine()
                sb2.AppendLine("Mappings:")
                sb2.AppendLine(String.Join(vbCrLf, mapping))
            End If
            If errorList.Count > 0 Then
                sb.AppendLine()
                sb.AppendLine("Errors:")
                sb.AppendLine(String.Join(vbCrLf, errorList))
                sb2.AppendLine()
                sb2.AppendLine("Errors:")
                sb2.AppendLine(String.Join(vbCrLf, errorList))
            End If

            Dim msg As String = ""
            If errorList.Count > 0 OrElse mapping.Count > 0 Then
                msg = "AI Rename results:" & vbCrLf & sb2.ToString()
                SLib.PutInClipboard(msg)
                msg = "AI Rename completed and result in the clipboard." & vbCrLf & vbCrLf & sb.ToString()
            Else
                msg = "AI Rename completed." & vbCrLf & vbCrLf & sb.ToString()
            End If
            ShowCustomMessageBox(msg)

        Catch ex As Exception
            ShowCustomMessageBox("AI Rename failed: " & ex.Message)
        Finally
            If do2ndModel AndAlso originalConfigLoaded Then
                Try
                    RestoreDefaults(_context, originalConfig)
                Catch
                End Try
                originalConfigLoaded = False
            End If
        End Try
    End Sub

    ''' <summary>
    ''' Truncates text to maxLen characters; appends marker if truncated.
    ''' </summary>
    ''' <param name="text">Input text to evaluate.</param>
    ''' <param name="maxLen">Maximum allowed length.</param>
    ''' <returns>Original text if within limit; truncated text with marker otherwise.</returns>
    Private Function TruncateForPrompt(text As String, maxLen As Integer) As String
        If String.IsNullOrEmpty(text) Then Return ""
        If text.Length <= maxLen Then Return text
        Return text.Substring(0, maxLen) & vbCrLf & "...[TRUNCATED]"
    End Function

    ''' <summary>
    ''' Removes invalid filename characters, trims, collapses whitespace, and ensures non-empty.
    ''' </summary>
    ''' <param name="rawName">Raw filename candidate.</param>
    ''' <returns>Sanitized filename body (without extension).</returns>
    Private Function SanitizeFileName(rawName As String) As String
        Dim invalid = Path.GetInvalidFileNameChars()
        Dim cleaned = New String(rawName.Where(Function(ch) Not invalid.Contains(ch)).ToArray())
        cleaned = cleaned.Trim()
        ' Remove dangerous trailing dots/spaces
        While cleaned.EndsWith(".") OrElse cleaned.EndsWith(" ")
            cleaned = cleaned.Substring(0, cleaned.Length - 1)
        End While
        ' Fallback if empty
        If String.IsNullOrWhiteSpace(cleaned) Then cleaned = "Unnamed"
        ' Optional: collapse whitespace
        cleaned = System.Text.RegularExpressions.Regex.Replace(cleaned, "\s+", " ")
        Return cleaned
    End Function

    ''' <summary>
    ''' Builds a unique full path by appending a numeric suffix if the target already exists.
    ''' </summary>
    ''' <param name="dir">Target directory.</param>
    ''' <param name="baseName">Filename body without extension.</param>
    ''' <param name="extension">Original file extension (including dot).</param>
    ''' <returns>Unique path in the specified directory.</returns>
    Private Function BuildUniqueTargetPath(dir As String, baseName As String, extension As String) As String
        Dim candidate = Path.Combine(dir, baseName & extension)
        If Not File.Exists(candidate) Then Return candidate
        Dim i As Integer = 2
        While True
            Dim nextCandidate = Path.Combine(dir, baseName & " (" & i & ")" & extension)
            If Not File.Exists(nextCandidate) Then Return nextCandidate
            i += 1
        End While
    End Function

End Class