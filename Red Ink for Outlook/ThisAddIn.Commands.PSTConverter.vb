' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.
' For license to use see https://redink.ai.
'
' =============================================================================
' File: ThisAddIn.Commands.PSTConverter.vb
' Purpose:
'   Exports Outlook/PST mail content to UTF-8 text files, including attachments.
'
' Architecture:
'  - Source selection:
'      * User chooses between PST file and Outlook folder source.
'      * PST files are mounted temporarily and removed again after export.
'      * Outlook source can be current folder or a folder chosen via Outlook picker.
'  - Scope options:
'      * Optional recursion into Outlook subfolders.
'      * Optional mirroring of Outlook subfolders on disk.
'      * Attachments exported either inline or as separate .txt files.
'  - Mail export:
'      * Each mail becomes MAIL000000.txt (UTF-8).
'      * Optional separate attachment files: MAIL000000-ATT000.txt.
'      * Embedded Outlook items (.msg/.eml / olEmbeddeditem) are processed recursively up to 5 levels.
'  - Logging:
'      * index.csv is always created.
'      * Unsupported/failed attachment extraction creates placeholders and a separate error log.
'  - Progress / cancellation:
'      * Uses ProgressBarModule and supports cancellation.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports Microsoft.Office.Interop.Outlook
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods
Imports SLib = SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    Private Const PSTExport_MaxEmbeddedDepth As Integer = 5

    Private NotInheritable Class PstExportSource
        Public Property RootFolder As MAPIFolder
        Public Property SourceDisplayName As String
        Public Property AddedStore As Store
        Public Property IsWholePstRoot As Boolean
    End Class

    Private NotInheritable Class PstExportOptions
        Public Property RecurseFolders As Boolean
        Public Property MirrorFoldersOnDisk As Boolean
        Public Property InlineAttachments As Boolean
        Public Property OutputRootDirectory As String
        Public Property EnablePdfOcr As Boolean
        Public Property EnableBinaryMediaExtraction As Boolean
    End Class

    Private NotInheritable Class PstExportState
        Public Property NextItemNumber As Integer = 0
        Public Property ExportedItemCount As Integer = 0
        Public Property FailedItemCount As Integer = 0
        Public Property PlaceholderCount As Integer = 0
        Public Property Cancelled As Boolean = False

        Public ReadOnly Property ErrorLines As New List(Of String)()
        Public ReadOnly Property IndexLines As New List(Of String)()
    End Class

    Private NotInheritable Class PstItemSnapshot
        Public Property EntryID As String
        Public Property StoreID As String
        Public Property SortKey As String
        Public Property SubjectPreview As String
    End Class

    ''' <summary>
    ''' Exports mails from a PST file or an Outlook folder to UTF-8 text files,
    ''' including attachment text extraction, placeholders, index.csv and error log.
    ''' </summary>
    Public Async Sub ExportPstContentToText()

        Dim source As PstExportSource = Nothing
        Dim options As PstExportOptions = Nothing
        Dim state As New PstExportState()
        Dim totalItemCount As Integer = 0

        Try
            source = PromptForPstExportSource()
            If source Is Nothing OrElse source.RootFolder Is Nothing Then
                Return
            End If

            options = PromptForPstExportOptions(source)
            If options Is Nothing OrElse String.IsNullOrWhiteSpace(options.OutputRootDirectory) Then
                Return
            End If

            If source.IsWholePstRoot Then
                options.RecurseFolders = True
            End If

            Try
                If Not Directory.Exists(options.OutputRootDirectory) Then
                    Directory.CreateDirectory(options.OutputRootDirectory)
                End If
            Catch ex As System.Exception
                ShowCustomMessageBox($"Failed to create output directory: {ex.Message}", AN)
                Return
            End Try

            totalItemCount = CountExportableItems(source.RootFolder, options.RecurseFolders)

            If totalItemCount <= 0 Then
                ShowCustomMessageBox("No exportable Outlook items were found in the selected source.", AN)
                Return
            End If

            state.IndexLines.Add(CsvLine(
            "ItemId",
            "RelativePath",
            "OutlookFolderPath",
            "ItemType",
            "MessageClass",
            "Subject",
            "Participants",
            "PrimaryDate",
            "AttachmentCount",
            "AttachmentOutputs",
            "BodyLength",
            "Status"))

            ProgressBarModule.GlobalProgressValue = 0
            ProgressBarModule.GlobalProgressMax = totalItemCount
            ProgressBarModule.GlobalProgressLabel = "Initializing..."
            ProgressBarModule.CancelOperation = False
            ProgressBarModule.ShowProgressBarInSeparateThread(AN & " PST / Outlook Export", "Starting export...")

            Try
                Await ExportFolderRecursiveAsync(source.RootFolder, options, state, "", totalItemCount)
            Finally
                ProgressBarModule.CancelOperation = True
            End Try

            WriteExportLogs(options, state)

            Dim summary As New StringBuilder()

            If state.Cancelled Then
                summary.AppendLine("Operation was cancelled by user.")
                summary.AppendLine()
            End If

            summary.AppendLine($"Exported items: {state.ExportedItemCount}")
            summary.AppendLine($"Failed items: {state.FailedItemCount}")
            summary.AppendLine($"Attachment placeholders: {state.PlaceholderCount}")
            summary.AppendLine($"Output directory: {options.OutputRootDirectory}")
            summary.AppendLine($"Index: {Path.Combine(options.OutputRootDirectory, "index.csv")}")
            summary.AppendLine($"Error log: {Path.Combine(options.OutputRootDirectory, "errors.log")}")

            ShowCustomMessageBox(summary.ToString().TrimEnd(), AN & " PST / Outlook Export")

        Catch ex As System.Exception
            ProgressBarModule.CancelOperation = True
            ShowCustomMessageBox($"Export failed: {ex.Message}", AN)
        Finally
            CleanupMountedStore(source)
        End Try

    End Sub

    Private Function PromptForPstExportOptions(source As PstExportSource) As PstExportOptions

        Dim recurseChoice As Integer = ShowCustomYesNoBox(
        "Do you want to include Outlook subfolders?",
        "Yes, include subfolders",
        "No, only the selected folder",
        AN & " PST / Outlook Export")

        If recurseChoice = 0 Then
            Return Nothing
        End If

        Dim recurseFolders As Boolean = (recurseChoice = 1)
        Dim mirrorFoldersOnDisk As Boolean = False

        If recurseFolders Then
            Dim mirrorChoice As Integer = ShowCustomYesNoBox(
            "Do you want Outlook subfolders to be created as subfolders on disk as well?",
            "Yes, mirror Outlook folders on disk",
            "No, keep one flat output directory",
            AN & " PST / Outlook Export")

            If mirrorChoice = 0 Then
                Return Nothing
            End If

            mirrorFoldersOnDisk = (mirrorChoice = 1)
        End If

        Dim attachmentChoice As Integer = ShowCustomYesNoBox(
        "How should attachments be exported?",
        "Inline in the main item file",
        "As separate attachment text files",
        AN & " PST / Outlook Export")

        If attachmentChoice = 0 Then
            Return Nothing
        End If

        Dim inlineAttachments As Boolean = (attachmentChoice = 1)

        Dim enablePdfOcr As Boolean = False
        If SharedMethods.IsOcrAvailable(_context) Then
            Dim ocrChoice As Integer = ShowCustomYesNoBox(
            "Do you want OCR enabled for PDF attachments when needed?",
            "Yes, enable PDF OCR",
            "No, keep direct text extraction only",
            AN & " PST / Outlook Export")

            If ocrChoice = 0 Then
                Return Nothing
            End If

            enablePdfOcr = (ocrChoice = 1)
        End If

        Dim enableBinaryMediaExtraction As Boolean = False
        Dim mediaChoice As Integer = ShowCustomYesNoBox(
        "Do you want AI extraction for image/audio/video attachments if your configured model supports them?",
        "Yes, enable binary/media extraction",
        "No, create placeholders for such files",
        AN & " PST / Outlook Export")

        If mediaChoice = 0 Then
            Return Nothing
        End If

        enableBinaryMediaExtraction = (mediaChoice = 1)

        Dim parentDirectory As String = Nothing
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select the parent directory for the export"
            fbd.ShowNewFolderButton = True

            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            If Not String.IsNullOrWhiteSpace(desktopPath) AndAlso Directory.Exists(desktopPath) Then
                fbd.SelectedPath = desktopPath
            End If

            Dim __safeDialogOwner255 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner255 IsNot Nothing, fbd.ShowDialog(__safeDialogOwner255), fbd.ShowDialog()) <> DialogResult.OK Then
                Return Nothing
            End If

            parentDirectory = fbd.SelectedPath
        End Using

        If String.IsNullOrWhiteSpace(parentDirectory) Then
            Return Nothing
        End If

        Dim outputRoot As String = CreateUniqueExportDirectory(parentDirectory, source.SourceDisplayName)

        Return New PstExportOptions() With {
        .RecurseFolders = recurseFolders,
        .MirrorFoldersOnDisk = mirrorFoldersOnDisk,
        .InlineAttachments = inlineAttachments,
        .OutputRootDirectory = outputRoot,
        .EnablePdfOcr = enablePdfOcr,
        .EnableBinaryMediaExtraction = enableBinaryMediaExtraction
    }

    End Function


    Private Function PromptForPstExportSource() As PstExportSource

        Dim outlookApp As Microsoft.Office.Interop.Outlook.Application = Globals.ThisAddIn.Application
        If outlookApp Is Nothing Then
            ShowCustomMessageBox("Outlook is not available.", AN)
            Return Nothing
        End If

        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Nothing
        Try
            ns = outlookApp.GetNamespace("MAPI")
        Catch ex As System.Exception
            ShowCustomMessageBox($"Could not access Outlook MAPI namespace: {ex.Message}", AN)
            Return Nothing
        End Try

        Dim sourceChoice As Integer = ShowCustomYesNoBox(
        "What do you want to export?",
        "PST file",
        "Outlook folder",
        AN & " PST / Outlook Export")

        If sourceChoice = 0 Then
            Return Nothing
        End If

        If sourceChoice = 1 Then
            Return PromptForPstFileSource(ns)
        End If

        Return PromptForOutlookFolderSource(outlookApp, ns)

    End Function


    Private Function PromptForPstFileSource(ns As Microsoft.Office.Interop.Outlook.NameSpace) As PstExportSource

        Dim pstPath As String = Nothing

        Using ofd As New OpenFileDialog()
            ofd.Filter = "Outlook Data Files (*.pst)|*.pst|All Files (*.*)|*.*"
            ofd.Title = "Select a PST File"
            ofd.Multiselect = False

            Dim __safeDialogOwner324 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner324 IsNot Nothing, ofd.ShowDialog(__safeDialogOwner324), ofd.ShowDialog()) <> DialogResult.OK Then
                Return Nothing
            End If

            pstPath = ofd.FileName
        End Using

        If String.IsNullOrWhiteSpace(pstPath) OrElse Not File.Exists(pstPath) Then
            ShowCustomMessageBox("The selected PST file does not exist.", AN)
            Return Nothing
        End If

        Dim addedStore As Store = Nothing

        Try
            ns.AddStore(pstPath)
            addedStore = FindStoreByFilePath(ns, pstPath)

            If addedStore Is Nothing Then
                ShowCustomMessageBox("The PST file was added, but its Outlook store could not be found.", AN)
                Return Nothing
            End If

            Dim rootFolder As MAPIFolder = Nothing
            Try
                rootFolder = addedStore.GetRootFolder()
            Catch
            End Try

            If rootFolder Is Nothing Then
                ShowCustomMessageBox("The PST root folder could not be accessed.", AN)
                CleanupMountedStore(New PstExportSource() With {.AddedStore = addedStore})
                Return Nothing
            End If

            Dim scopeChoice As Integer = ShowCustomYesNoBox(
            "Do you want to export the whole PST or choose a folder inside that PST?" & vbCrLf & vbCrLf &
            "If you choose the whole PST, all subfolders will be included automatically.",
            "Whole PST",
            "Choose folder",
            AN & " PST / Outlook Export")

            If scopeChoice = 0 Then
                CleanupMountedStore(New PstExportSource() With {.AddedStore = addedStore})
                Return Nothing
            End If

            Dim selectedFolder As MAPIFolder = rootFolder
            Dim isWholePstRoot As Boolean = True

            If scopeChoice = 2 Then
                selectedFolder = ns.PickFolder()
                isWholePstRoot = False

                If selectedFolder Is Nothing Then
                    CleanupMountedStore(New PstExportSource() With {.AddedStore = addedStore})
                    Return Nothing
                End If

                If Not FolderBelongsToStore(selectedFolder, addedStore) Then
                    ShowCustomMessageBox("The selected folder does not belong to the chosen PST. Please try again.", AN)
                    CleanupMountedStore(New PstExportSource() With {.AddedStore = addedStore})
                    Return Nothing
                End If
            End If

            Dim displayName As String = Path.GetFileNameWithoutExtension(pstPath)
            If selectedFolder IsNot Nothing Then
                Try
                    displayName &= "_" & SanitizePathSegment(selectedFolder.Name)
                Catch
                End Try
            End If

            Return New PstExportSource() With {
            .RootFolder = selectedFolder,
            .SourceDisplayName = displayName,
            .AddedStore = addedStore,
            .IsWholePstRoot = isWholePstRoot
        }

        Catch ex As System.Exception
            CleanupMountedStore(New PstExportSource() With {.AddedStore = addedStore})
            ShowCustomMessageBox($"Could not mount the PST file: {ex.Message}", AN)
            Return Nothing
        End Try

    End Function

    Private Function PromptForOutlookFolderSource(
        outlookApp As Microsoft.Office.Interop.Outlook.Application,
        ns As Microsoft.Office.Interop.Outlook.NameSpace) As PstExportSource

        Dim useCurrentChoice As Integer = ShowCustomYesNoBox(
            "Do you want to use the current Outlook folder or choose another folder?",
            "Use current folder",
            "Choose folder",
            AN & " PST / Outlook Export")

        If useCurrentChoice = 0 Then
            Return Nothing
        End If

        Dim selectedFolder As MAPIFolder = Nothing

        If useCurrentChoice = 1 Then
            Dim explorer As Explorer = Nothing
            Try
                explorer = outlookApp.ActiveExplorer()
            Catch
            End Try

            If explorer Is Nothing Then
                ShowCustomMessageBox("No active Outlook explorer window was found.", AN)
                Return Nothing
            End If

            Try
                selectedFolder = explorer.CurrentFolder
            Catch
            End Try
        Else
            selectedFolder = ns.PickFolder()
        End If

        If selectedFolder Is Nothing Then
            ShowCustomMessageBox("No Outlook folder was selected.", AN)
            Return Nothing
        End If

        Dim displayName As String = "Outlook_" & SanitizePathSegment(selectedFolder.Name)

        Return New PstExportSource() With {
            .RootFolder = selectedFolder,
            .SourceDisplayName = displayName,
            .AddedStore = Nothing
        }

    End Function


    Private Async Function ExportFolderRecursiveAsync(
    folder As MAPIFolder,
    options As PstExportOptions,
    state As PstExportState,
    relativeDiskPath As String,
    totalItemCount As Integer) As Task

        If folder Is Nothing OrElse state Is Nothing OrElse options Is Nothing Then Return

        If ProgressBarModule.CancelOperation Then
            state.Cancelled = True
            Return
        End If

        Dim targetDirectory As String = options.OutputRootDirectory
        If options.MirrorFoldersOnDisk AndAlso Not String.IsNullOrWhiteSpace(relativeDiskPath) Then
            targetDirectory = Path.Combine(options.OutputRootDirectory, relativeDiskPath)
        End If

        If Not Directory.Exists(targetDirectory) Then
            Directory.CreateDirectory(targetDirectory)
        End If

        Dim snapshots As List(Of PstItemSnapshot) = GetSortedItemSnapshots(folder)
        Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")

        For Each snapshot In snapshots
            System.Windows.Forms.Application.DoEvents()

            If ProgressBarModule.CancelOperation Then
                state.Cancelled = True
                Return
            End If

            Dim itemObj As Object = Nothing

            Try
                ProgressBarModule.GlobalProgressValue = state.ExportedItemCount + state.FailedItemCount
                ProgressBarModule.GlobalProgressLabel =
                $"Exporting item {Math.Min(ProgressBarModule.GlobalProgressValue + 1, totalItemCount)} of {totalItemCount}: {snapshot.SubjectPreview}"

                itemObj = ns.GetItemFromID(snapshot.EntryID, snapshot.StoreID)
                If itemObj Is Nothing Then
                    state.FailedItemCount += 1
                    AppendError(state, $"ITEM-LOAD|EntryID={snapshot.EntryID}|Reason=GetItemFromID returned Nothing")
                    Continue For
                End If

                Await ExportSingleOutlookItemAsync(itemObj, folder, targetDirectory, options, state)

            Catch ex As System.Exception
                state.FailedItemCount += 1
                AppendError(state, $"ITEM-ERROR|EntryID={snapshot.EntryID}|Folder={TryGetFolderPath(folder)}|Message={ex.Message}")
            Finally
                ReleaseComObjectIfNeeded(itemObj)
            End Try
        Next

        If Not options.RecurseFolders Then
            Return
        End If

        For Each subFolder As MAPIFolder In GetSortedSubFolders(folder)
            If subFolder Is Nothing Then Continue For

            Dim childRelativeDiskPath As String = relativeDiskPath
            If options.MirrorFoldersOnDisk Then
                childRelativeDiskPath = AppendRelativeDiskPath(relativeDiskPath, SanitizePathSegment(subFolder.Name))
            End If

            Await ExportFolderRecursiveAsync(subFolder, options, state, childRelativeDiskPath, totalItemCount)
        Next

    End Function

    Private Async Function ExportSingleOutlookItemAsync(
    item As Object,
    owningFolder As MAPIFolder,
    outputDirectory As String,
    options As PstExportOptions,
    state As PstExportState) As Task

        If item Is Nothing Then Return

        state.NextItemNumber += 1
        Dim itemId As String = "MAIL" & state.NextItemNumber.ToString("000000")
        Dim itemFileName As String = itemId & ".txt"
        Dim itemPath As String = Path.Combine(outputDirectory, itemFileName)

        Dim sb As New StringBuilder()
        Dim attachmentOutputNames As New List(Of String)()
        Dim localPlaceholderCountBefore As Integer = state.PlaceholderCount

        AppendOutlookItemHeader(sb, item, owningFolder, itemId)

        Dim bodyText As String = GetOutlookItemTextBody(item)

        sb.AppendLine("=== BODY / PRIMARY TEXT ===")
        sb.AppendLine(bodyText)

        Dim detailsText As String = BuildItemSpecificDetailsText(item)
        If Not String.IsNullOrWhiteSpace(detailsText) Then
            sb.AppendLine()
            sb.AppendLine("=== STRUCTURED DETAILS ===")
            sb.AppendLine(detailsText.Trim())
        End If

        Dim attachments As Attachments = TryGetAttachments(item)
        Dim attachmentCount As Integer = If(attachments Is Nothing, 0, attachments.Count)

        sb.AppendLine()
        sb.AppendLine("=== ATTACHMENTS ===")
        sb.AppendLine($"Count: {attachmentCount}")

        If attachmentCount > 0 Then
            For i As Integer = 1 To attachmentCount
                System.Windows.Forms.Application.DoEvents()

                If ProgressBarModule.CancelOperation Then
                    state.Cancelled = True
                    Exit For
                End If

                Dim attachment As Attachment = Nothing

                Try
                    attachment = attachments(i)
                    Dim attachmentId As String = $"{itemId}-ATT{i:000}"

                    If options.InlineAttachments Then
                        sb.AppendLine()
                        sb.AppendLine(New String("="c, 80))
                        sb.AppendLine($"ATTACHMENT {i:000}")
                        sb.AppendLine(New String("="c, 80))
                        sb.AppendLine(Await RenderAttachmentTextAsync(attachment, attachmentId, state, options, 1))
                    Else
                        Dim attachmentFileName As String = attachmentId & ".txt"
                        Dim attachmentFilePath As String = Path.Combine(outputDirectory, attachmentFileName)

                        Dim attachmentText As String =
                        Await RenderAttachmentTextAsync(attachment, attachmentId, state, options, 1)

                        File.WriteAllText(attachmentFilePath, attachmentText, New UTF8Encoding(False))
                        attachmentOutputNames.Add(attachmentFileName)

                        sb.AppendLine($"Attachment {i:000}: {TryGetAttachmentFileName(attachment)} -> {attachmentFileName}")
                    End If

                Catch ex As System.Exception
                    state.PlaceholderCount += 1
                    AppendError(state, $"ATTACHMENT-ERROR|ItemId={itemId}|AttachmentIndex={i}|Message={ex.Message}")
                    sb.AppendLine($"Attachment {i:000}: [placeholder created due to error]")
                Finally
                    ReleaseComObjectIfNeeded(attachment)
                End Try
            Next
        End If

        File.WriteAllText(itemPath, sb.ToString(), New UTF8Encoding(False))

        Dim status As String = If(state.PlaceholderCount > localPlaceholderCountBefore, "Partial", "OK")

        state.IndexLines.Add(CsvLine(
        itemId,
        MakeRelativePath(options.OutputRootDirectory, itemPath),
        TryGetFolderPath(owningFolder),
        GetOutlookItemDisplayTypeName(item),
        TryGetOutlookItemMessageClass(item),
        GetOutlookItemSubject(item),
        GetOutlookItemParticipants(item),
        FormatNullableDate(TryGetOutlookItemPrimaryDate(item)),
        attachmentCount.ToString(),
        String.Join(";", attachmentOutputNames),
        bodyText.Length.ToString(),
        status))

        state.ExportedItemCount += 1

    End Function

    Private Async Function RenderAttachmentTextAsync(
    attachment As Attachment,
    attachmentId As String,
    state As PstExportState,
    options As PstExportOptions,
    depth As Integer) As Task(Of String)

        Dim sb As New StringBuilder()

        Dim attachmentFileName As String = TryGetAttachmentFileName(attachment)
        Dim attachmentExtension As String = ""
        Try
            attachmentExtension = Path.GetExtension(attachmentFileName).ToLowerInvariant()
        Catch
            attachmentExtension = ""
        End Try

        sb.AppendLine($"Attachment ID: {attachmentId}")
        sb.AppendLine($"File name: {attachmentFileName}")
        sb.AppendLine($"Original type: {TryGetAttachmentTypeDescription(attachment)}")
        sb.AppendLine($"Recursion depth: {depth}")

        If depth > PSTExport_MaxEmbeddedDepth Then
            state.PlaceholderCount += 1
            Dim placeholder As String = BuildAttachmentPlaceholderText(
            attachmentFileName,
            TryGetAttachmentTypeDescription(attachment),
            $"Maximum embedded Outlook recursion depth ({PSTExport_MaxEmbeddedDepth}) reached.")
            sb.AppendLine()
            sb.AppendLine(placeholder)
            AppendError(state, $"ATTACHMENT-DEPTH|AttachmentId={attachmentId}|File={attachmentFileName}|Reason=Maximum depth reached")
            Return sb.ToString().TrimEnd()
        End If

        If IsEmbeddedOutlookAttachment(attachment) Then
            sb.AppendLine()
            sb.AppendLine(Await RenderEmbeddedOutlookAttachmentTextAsync(attachment, attachmentId, state, options, depth))
            Return sb.ToString().TrimEnd()
        End If

        Dim tempPath As String = Nothing

        Try
            tempPath = SaveAttachmentToTemporaryFile(attachment, attachmentFileName, False)

            If String.IsNullOrWhiteSpace(tempPath) OrElse Not File.Exists(tempPath) Then
                state.PlaceholderCount += 1
                Dim placeholder As String = BuildAttachmentPlaceholderText(
                attachmentFileName,
                TryGetAttachmentTypeDescription(attachment),
                "Attachment could not be saved to a temporary file.")
                sb.AppendLine()
                sb.AppendLine(placeholder)
                AppendError(state, $"ATTACHMENT-SAVE|AttachmentId={attachmentId}|File={attachmentFileName}|Reason=Temporary save failed")
                Return sb.ToString().TrimEnd()
            End If

            Dim extractedText As String = Await ExtractTextFromSavedAttachmentAsync(tempPath, options)

            If IsExtractionFailure(extractedText) Then
                state.PlaceholderCount += 1
                Dim placeholder As String = BuildAttachmentPlaceholderText(
                attachmentFileName,
                TryGetAttachmentTypeDescription(attachment),
                If(String.IsNullOrWhiteSpace(extractedText), "No text could be extracted.", extractedText))
                sb.AppendLine()
                sb.AppendLine(placeholder)
                AppendError(state, $"ATTACHMENT-EXTRACT|AttachmentId={attachmentId}|File={attachmentFileName}|Reason={SafeSingleLine(If(extractedText, "No text could be extracted."), 200)}")
                Return sb.ToString().TrimEnd()
            End If

            sb.AppendLine()
            sb.AppendLine("=== EXTRACTED TEXT ===")
            sb.AppendLine(extractedText.Trim())

            Return sb.ToString().TrimEnd()

        Catch ex As System.Exception
            state.PlaceholderCount += 1
            Dim placeholder As String = BuildAttachmentPlaceholderText(
            attachmentFileName,
            TryGetAttachmentTypeDescription(attachment),
            ex.Message)
            sb.AppendLine()
            sb.AppendLine(placeholder)
            AppendError(state, $"ATTACHMENT-ERROR|AttachmentId={attachmentId}|File={attachmentFileName}|Message={ex.Message}")
            Return sb.ToString().TrimEnd()
        Finally
            Try
                If Not String.IsNullOrWhiteSpace(tempPath) AndAlso File.Exists(tempPath) Then
                    File.Delete(tempPath)
                End If
            Catch
            End Try
        End Try

    End Function


    Private Function ReadMsgAttachmentText(filePath As String) As String
        Return SharedMethods.ReadMsgSandboxed(
            filePath,
            Function(msgPath As String, tempDir As String, ByRef nestedFiles As List(Of String)) As String
                nestedFiles = New List(Of String)()

                Dim sharedItem As Object = Nothing
                Dim attachments As Attachments = Nothing

                Try
                    Dim session As Microsoft.Office.Interop.Outlook.NameSpace =
                        Globals.ThisAddIn.Application.GetNamespace("MAPI")

                    sharedItem = session.OpenSharedItem(msgPath)
                    If sharedItem Is Nothing Then Return Nothing

                    Dim sb As New StringBuilder()

                    AppendOutlookItemHeader(
                        sb,
                        sharedItem,
                        Nothing,
                        Path.GetFileNameWithoutExtension(msgPath))

                    Dim bodyText As String = GetOutlookItemTextBody(sharedItem)
                    sb.AppendLine("=== BODY / PRIMARY TEXT ===")
                    sb.AppendLine(bodyText)

                    Dim detailsText As String = BuildItemSpecificDetailsText(sharedItem)
                    If Not String.IsNullOrWhiteSpace(detailsText) Then
                        sb.AppendLine()
                        sb.AppendLine("=== STRUCTURED DETAILS ===")
                        sb.AppendLine(detailsText.Trim())
                    End If

                    attachments = TryGetAttachments(sharedItem)
                    Dim attachmentCount As Integer = If(attachments Is Nothing, 0, attachments.Count)

                    sb.AppendLine()
                    sb.AppendLine("=== ATTACHMENTS ===")
                    sb.AppendLine($"Count: {attachmentCount}")

                    If attachmentCount > 0 Then
                        For i As Integer = 1 To attachmentCount
                            Dim nestedAttachment As Attachment = Nothing

                            Try
                                nestedAttachment = attachments(i)

                                Dim nestedName As String = TryGetAttachmentFileName(nestedAttachment)
                                If String.IsNullOrWhiteSpace(nestedName) Then
                                    nestedName = $"attachment_{i:000}"
                                End If

                                If nestedAttachment.Type = OlAttachmentType.olEmbeddeditem AndAlso
                                   Not nestedName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) AndAlso
                                   Not nestedName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase) Then
                                    nestedName = Path.GetFileNameWithoutExtension(nestedName) & ".msg"
                                End If

                                Dim nestedPath As String = Path.Combine(tempDir, nestedName)
                                Dim counter As Integer = 1

                                While File.Exists(nestedPath)
                                    nestedPath = Path.Combine(
                                        tempDir,
                                        Path.GetFileNameWithoutExtension(nestedName) &
                                        $"_{counter}" &
                                        Path.GetExtension(nestedName))
                                    counter += 1
                                End While

                                nestedAttachment.SaveAsFile(nestedPath)
                                nestedFiles.Add(nestedPath)

                                sb.AppendLine($"  - {nestedName}")
                            Catch ex As System.Exception
                                sb.AppendLine($"  - [attachment {i:000} could not be extracted: {ex.Message}]")
                            Finally
                                ReleaseComObjectIfNeeded(nestedAttachment)
                            End Try
                        Next
                    End If

                    Return sb.ToString().TrimEnd()
                Catch
                    Return Nothing
                Finally
                    ReleaseComObjectIfNeeded(attachments)
                    ReleaseComObjectIfNeeded(sharedItem)
                End Try
            End Function)
    End Function

    Private Async Function ExtractTextFromSavedAttachmentAsync(
    filePath As String,
    options As PstExportOptions) As Task(Of String)

        If String.IsNullOrWhiteSpace(filePath) OrElse Not File.Exists(filePath) Then
            Return "Error: temporary attachment file does not exist."
        End If

        Dim extracted As String = Nothing
        Dim label As String = Nothing
        Dim ext As String = Path.GetExtension(filePath).ToLowerInvariant()

        Select Case ext
            Case ".msg"
                Dim msgText As String = ReadMsgAttachmentText(filePath)
                If Not String.IsNullOrWhiteSpace(msgText) AndAlso
                   Not msgText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return msgText
                End If

            Case ".eml"
                Dim emlText As String = SharedMethods.ReadEmlSandboxed(filePath)
                If Not String.IsNullOrWhiteSpace(emlText) AndAlso
                   Not emlText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return emlText
                End If

            Case ".rtf"
                Dim rtfText As String = SharedMethods.ReadRtfAsText(filePath, ReturnErrorInsteadOfEmpty:=True)
                If Not String.IsNullOrWhiteSpace(rtfText) AndAlso
                   Not rtfText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return rtfText
                End If

            Case ".docx", ".docm"
                Dim docText As String = SharedMethods.ReadDocxSandboxed(filePath)
                If Not String.IsNullOrWhiteSpace(docText) AndAlso
                   Not docText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return docText
                End If

            Case ".xlsx", ".xlsm"
                Dim xlsxText As String = SharedMethods.ReadXlsxSandboxed(filePath, silent:=True, askWorksheetSelection:=False)
                If Not String.IsNullOrWhiteSpace(xlsxText) AndAlso
                   Not xlsxText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return xlsxText
                End If

            Case ".pptx", ".pptm"
                Dim pptText As String = SharedMethods.ReadPptxSandboxed(filePath)
                If Not String.IsNullOrWhiteSpace(pptText) AndAlso
                   Not pptText.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) Then
                    Return pptText
                End If
        End Select

        Try
            If TryExtractOfficeText(filePath, extracted, label) Then
                If Not String.IsNullOrWhiteSpace(extracted) Then
                    Return extracted
                End If

                Return $"Error: safe Office extraction returned empty content for '{ext}'."
            End If
        Catch ex As System.Exception
            Return $"Error: Office extraction failed for '{ext}': {ex.Message}"
        End Try

        Try
            If TryExtractTextLike(filePath, extracted, label) Then
                If Not String.IsNullOrWhiteSpace(extracted) Then
                    Return extracted
                End If

                Return $"Error: text-like extraction returned empty content for '{ext}'."
            End If
        Catch ex As System.Exception
            Return $"Error: text-like extraction failed for '{ext}': {ex.Message}"
        End Try

        If ext = ".pdf" Then
            Try
                Dim pdfText As String =
                Await SharedMethods.ReadPdfAsText(
                    filePath,
                    ReturnErrorInsteadOfEmpty:=True,
                    DoOCR:=options.EnablePdfOcr,
                    AskUser:=False,
                    context:=_context)

                If String.IsNullOrWhiteSpace(pdfText) Then
                    Return "Error: PDF extraction returned empty content."
                End If

                Return pdfText
            Catch ex As System.Exception
                Return $"Error: PDF extraction failed: {ex.Message}"
            End Try
        End If

        Dim taskFlag As String = SharedMethods.TaskFlagForExtension(ext)
        If options.EnableBinaryMediaExtraction AndAlso Not String.IsNullOrWhiteSpace(taskFlag) Then
            Try
                Dim mediaText As String =
                Await SharedMethods.ReadBinaryFileViaLLM(
                    filePath,
                    _context,
                    askUser:=False,
                    taskFlag:=taskFlag)

                If String.IsNullOrWhiteSpace(mediaText) Then
                    Return $"Error: binary/media extraction returned empty content for '{ext}'."
                End If

                Return mediaText
            Catch ex As System.Exception
                Return $"Error: binary/media extraction failed for '{ext}': {ex.Message}"
            End Try
        End If

        Return $"Error: unsupported attachment type '{ext}' or no extraction path is configured."
    End Function

    Private Async Function RenderEmbeddedOutlookAttachmentTextAsync(
    attachment As Attachment,
    attachmentId As String,
    state As PstExportState,
    options As PstExportOptions,
    depth As Integer) As Task(Of String)

        Dim sb As New StringBuilder()
        Dim tempPath As String = Nothing
        Dim sharedItem As Object = Nothing

        Try
            tempPath = SaveAttachmentToTemporaryFile(attachment, TryGetAttachmentFileName(attachment), True)

            If String.IsNullOrWhiteSpace(tempPath) OrElse Not File.Exists(tempPath) Then
                state.PlaceholderCount += 1
                Dim placeholder As String = BuildAttachmentPlaceholderText(
                TryGetAttachmentFileName(attachment),
                TryGetAttachmentTypeDescription(attachment),
                "Embedded Outlook item could not be saved.")
                AppendError(state, $"EMBEDDED-SAVE|AttachmentId={attachmentId}|File={TryGetAttachmentFileName(attachment)}|Reason=Temporary save failed")
                Return placeholder
            End If

            Dim session As Microsoft.Office.Interop.Outlook.NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")
            sharedItem = session.OpenSharedItem(tempPath)

            If sharedItem Is Nothing Then
                state.PlaceholderCount += 1
                Return BuildAttachmentPlaceholderText(
                TryGetAttachmentFileName(attachment),
                TryGetAttachmentTypeDescription(attachment),
                "Embedded Outlook item could not be opened.")
            End If

            sb.AppendLine("=== EMBEDDED OUTLOOK ITEM ===")
            sb.AppendLine(Await BuildEmbeddedOutlookItemTextAsync(sharedItem, attachmentId & "-MAIL", state, options, depth))

            Return sb.ToString().TrimEnd()

        Catch ex As System.Exception
            state.PlaceholderCount += 1
            AppendError(state, $"EMBEDDED-ERROR|AttachmentId={attachmentId}|Message={ex.Message}")
            Return BuildAttachmentPlaceholderText(
            TryGetAttachmentFileName(attachment),
            TryGetAttachmentTypeDescription(attachment),
            ex.Message)
        Finally
            ReleaseComObjectIfNeeded(sharedItem)

            Try
                If Not String.IsNullOrWhiteSpace(tempPath) AndAlso File.Exists(tempPath) Then
                    File.Delete(tempPath)
                End If
            Catch
            End Try
        End Try

    End Function
    Private Sub AppendMailHeader(
        sb As StringBuilder,
        mail As MailItem,
        owningFolder As MAPIFolder,
        mailId As String)

        sb.AppendLine($"Mail ID: {mailId}")
        sb.AppendLine($"Subject: {TryGetMailSubject(mail)}")
        sb.AppendLine($"From: {TryGetMailSender(mail)}")
        sb.AppendLine($"To: {TryGetMailTo(mail)}")
        sb.AppendLine($"CC: {TryGetMailCc(mail)}")
        sb.AppendLine($"Sent: {TryGetMailSentOn(mail)}")
        sb.AppendLine($"Received: {TryGetMailReceivedTime(mail)}")
        sb.AppendLine($"Outlook Folder: {If(owningFolder Is Nothing, "", TryGetFolderPath(owningFolder))}")

        Try
            sb.AppendLine($"Conversation Topic: {If(mail.ConversationTopic, "")}")
        Catch
            sb.AppendLine("Conversation Topic: ")
        End Try

        Try
            sb.AppendLine($"Importance: {mail.Importance}")
        Catch
            sb.AppendLine("Importance: ")
        End Try

        Try
            sb.AppendLine($"Categories: {If(mail.Categories, "")}")
        Catch
            sb.AppendLine("Categories: ")
        End Try

        sb.AppendLine()
    End Sub

    Private Function CountMailItems(folder As MAPIFolder, recurseFolders As Boolean) As Integer
        If folder Is Nothing Then Return 0

        Dim count As Integer = 0
        Dim items As Items = Nothing

        Try
            items = folder.Items
        Catch
            items = Nothing
        End Try

        If items IsNot Nothing Then
            Try
                For i As Integer = 1 To items.Count
                    Dim itemObj As Object = Nothing
                    Try
                        itemObj = items(i)
                        If TypeOf itemObj Is MailItem Then
                            count += 1
                        End If
                    Catch
                    Finally
                        ReleaseComObjectIfNeeded(itemObj)
                    End Try
                Next
            Catch
            End Try
        End If

        If recurseFolders Then
            Dim subFolders As Folders = Nothing
            Try
                subFolders = folder.Folders
            Catch
                subFolders = Nothing
            End Try

            If subFolders IsNot Nothing Then
                For Each subFolderObj As Object In subFolders
                    Dim subFolder As MAPIFolder = TryCast(subFolderObj, MAPIFolder)
                    If subFolder Is Nothing Then Continue For

                    count += CountMailItems(subFolder, True)
                Next
            End If
        End If

        Return count
    End Function

    Private Function CreateUniqueExportDirectory(parentDirectory As String, sourceDisplayName As String) As String
        Dim baseName As String =
            "PSTExport_" &
            SanitizePathSegment(sourceDisplayName) &
            "_" &
            DateTime.Now.ToString("yyyyMMdd_HHmmss")

        Dim candidate As String = Path.Combine(parentDirectory, baseName)
        Dim suffix As Integer = 1

        While Directory.Exists(candidate)
            candidate = Path.Combine(parentDirectory, baseName & "_" & suffix.ToString())
            suffix += 1
        End While

        Return candidate
    End Function

    Private Function FindStoreByFilePath(ns As Microsoft.Office.Interop.Outlook.NameSpace, pstPath As String) As Store
        If ns Is Nothing OrElse String.IsNullOrWhiteSpace(pstPath) Then Return Nothing

        For Each storeObj As Object In ns.Stores
            Dim store As Store = TryCast(storeObj, Store)
            If store Is Nothing Then Continue For

            Try
                If String.Equals(store.FilePath, pstPath, StringComparison.OrdinalIgnoreCase) Then
                    Return store
                End If
            Catch
            End Try
        Next

        Return Nothing
    End Function

    Private Function FolderBelongsToStore(folder As MAPIFolder, store As Store) As Boolean
        If folder Is Nothing OrElse store Is Nothing Then Return False

        Try
            Return String.Equals(folder.Store.StoreID, store.StoreID, StringComparison.OrdinalIgnoreCase)
        Catch
            Return False
        End Try
    End Function

    Private Sub CleanupMountedStore(source As PstExportSource)
        If source Is Nothing OrElse source.AddedStore Is Nothing Then Return

        Try
            Dim ns As Microsoft.Office.Interop.Outlook.NameSpace = Globals.ThisAddIn.Application.GetNamespace("MAPI")
            Dim rootFolder As MAPIFolder = source.AddedStore.GetRootFolder()
            ns.RemoveStore(rootFolder)
        Catch
        End Try
    End Sub

    Private Function SaveAttachmentToTemporaryFile(
        attachment As Attachment,
        originalFileName As String,
        forceMsgExtension As Boolean) As String

        If attachment Is Nothing Then Return Nothing

        Dim safeName As String = SanitizeFileNameForFile(If(originalFileName, "attachment"))
        If String.IsNullOrWhiteSpace(safeName) Then
            safeName = "attachment"
        End If

        If forceMsgExtension Then
            If Not safeName.EndsWith(".msg", StringComparison.OrdinalIgnoreCase) AndAlso
               Not safeName.EndsWith(".eml", StringComparison.OrdinalIgnoreCase) Then
                safeName &= ".msg"
            End If
        End If

        Dim tempPath As String =
            Path.Combine(
                Path.GetTempPath(),
                $"{AN2}_pstexport_{Guid.NewGuid():N}_{safeName}")

        attachment.SaveAsFile(tempPath)
        Return tempPath
    End Function

    Private Function TryGetAttachmentFileName(attachment As Attachment) As String
        If attachment Is Nothing Then Return "attachment"

        Try
            Dim value As String = attachment.FileName
            If Not String.IsNullOrWhiteSpace(value) Then Return value
        Catch
        End Try

        Return "attachment"
    End Function

    Private Function TryGetAttachmentTypeDescription(attachment As Attachment) As String
        If attachment Is Nothing Then Return "Unknown"

        Dim typeName As String = "Unknown"
        Try
            typeName = attachment.Type.ToString()
        Catch
        End Try

        Dim ext As String = ""
        Try
            ext = Path.GetExtension(TryGetAttachmentFileName(attachment))
        Catch
        End Try

        If String.IsNullOrWhiteSpace(ext) Then
            Return typeName
        End If

        Return $"{typeName} ({ext})"
    End Function

    Private Function IsEmbeddedOutlookAttachment(attachment As Attachment) As Boolean
        If attachment Is Nothing Then Return False

        Try
            Return attachment.Type = OlAttachmentType.olEmbeddeditem
        Catch
            Return False
        End Try
    End Function


    Private Function IsExtractionFailure(text As String) As Boolean
        If String.IsNullOrWhiteSpace(text) Then
            Return True
        End If

        If text.StartsWith("Error:", StringComparison.OrdinalIgnoreCase) OrElse
       text.StartsWith("Error ", StringComparison.OrdinalIgnoreCase) OrElse
       text.StartsWith("Error", StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        Return False
    End Function

    Private Function BuildAttachmentPlaceholderText(fileName As String, originalType As String, reason As String) As String
        Dim sb As New StringBuilder()

        sb.AppendLine("=== PLACEHOLDER ===")
        sb.AppendLine($"File name: {If(fileName, "")}")
        sb.AppendLine($"Original type: {If(originalType, "")}")
        sb.AppendLine($"Reason: {If(reason, "Unsupported attachment or extraction failed.")}")

        Return sb.ToString().TrimEnd()
    End Function

    Private Sub WriteExportLogs(options As PstExportOptions, state As PstExportState)
        If options Is Nothing OrElse state Is Nothing Then Return

        Try
            Dim indexPath As String = Path.Combine(options.OutputRootDirectory, "index.csv")
            File.WriteAllLines(indexPath, state.IndexLines, New UTF8Encoding(False))
        Catch
        End Try

        Try
            Dim errorPath As String = Path.Combine(options.OutputRootDirectory, "errors.log")
            If state.ErrorLines.Count = 0 Then
                File.WriteAllText(errorPath, "No errors logged.", New UTF8Encoding(False))
            Else
                File.WriteAllLines(errorPath, state.ErrorLines, New UTF8Encoding(False))
            End If
        Catch
        End Try
    End Sub

    Private Sub AppendError(state As PstExportState, message As String)
        If state Is Nothing Then Return
        state.ErrorLines.Add($"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}")
    End Sub

    Private Function CsvLine(ParamArray values() As String) As String
        Dim escaped As New List(Of String)()

        For Each value As String In values
            Dim safeValue As String = If(value, "")
            safeValue = safeValue.Replace("""", """""")
            escaped.Add($"""{safeValue}""")
        Next

        Return String.Join(",", escaped)
    End Function

    Private Function MakeRelativePath(baseDirectory As String, fullPath As String) As String
        Try
            Dim baseUri As New Uri(AppendDirectorySeparator(baseDirectory))
            Dim fullUri As New Uri(fullPath)
            Return Uri.UnescapeDataString(baseUri.MakeRelativeUri(fullUri).ToString()).Replace("/"c, "\"c)
        Catch
            Return fullPath
        End Try
    End Function

    Private Function AppendDirectorySeparator(directoryPath As String) As String
        If String.IsNullOrWhiteSpace(directoryPath) Then Return directoryPath

        If directoryPath.EndsWith(IO.Path.DirectorySeparatorChar) OrElse
       directoryPath.EndsWith(IO.Path.AltDirectorySeparatorChar) Then
            Return directoryPath
        End If

        Return directoryPath & IO.Path.DirectorySeparatorChar
    End Function

    Private Function AppendRelativeDiskPath(currentPath As String, nextSegment As String) As String
        If String.IsNullOrWhiteSpace(currentPath) Then
            Return nextSegment
        End If

        Return Path.Combine(currentPath, nextSegment)
    End Function

    Private Function NormalizeText(text As String) As String
        If text Is Nothing Then Return String.Empty
        Return text.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Trim()
    End Function

    Private Function SanitizePathSegment(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "Folder"

        Dim result As String = name
        For Each ch As Char In Path.GetInvalidFileNameChars()
            result = result.Replace(ch, "_"c)
        Next

        result = result.Replace(":"c, "_"c).Trim()

        If String.IsNullOrWhiteSpace(result) Then
            result = "Folder"
        End If

        If result.Length > 80 Then
            result = result.Substring(0, 80)
        End If

        Return result
    End Function

    Private Function SanitizeFileNameForFile(name As String) As String
        If String.IsNullOrWhiteSpace(name) Then Return "file"

        Dim result As String = name
        For Each ch As Char In Path.GetInvalidFileNameChars()
            result = result.Replace(ch, "_"c)
        Next

        If String.IsNullOrWhiteSpace(result) Then
            result = "file"
        End If

        Return result
    End Function

    Private Function SafeSingleLine(text As String, maxLength As Integer) As String
        Dim value As String = If(text, "")
        value = value.Replace(vbCr, " ").Replace(vbLf, " ").Trim()

        If value.Length > maxLength Then
            value = value.Substring(0, maxLength) & "..."
        End If

        Return value
    End Function

    Private Function TryGetFolderPath(folder As MAPIFolder) As String
        If folder Is Nothing Then Return ""

        Try
            Return folder.FolderPath
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailSubject(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            Return If(mail.Subject, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailSender(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            Dim senderName As String = If(mail.SenderName, "")
            Dim senderEmail As String = ""
            Try
                senderEmail = If(mail.SenderEmailAddress, "")
            Catch
            End Try

            If String.IsNullOrWhiteSpace(senderEmail) Then
                Return senderName
            End If

            If String.IsNullOrWhiteSpace(senderName) Then
                Return senderEmail
            End If

            Return $"{senderName} <{senderEmail}>"
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailTo(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            Return If(mail.To, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailCc(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            Return If(mail.CC, "")
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailSentOn(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            If mail.SentOn = Date.MinValue Then Return ""
            Return mail.SentOn.ToString("yyyy-MM-dd HH:mm:ss")
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetMailReceivedTime(mail As MailItem) As String
        If mail Is Nothing Then Return ""
        Try
            If mail.ReceivedTime = Date.MinValue Then Return ""
            Return mail.ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
        Catch
            Return ""
        End Try
    End Function

    Private Sub ReleaseComObjectIfNeeded(obj As Object)
        If obj Is Nothing Then Return

        Try
            If Marshal.IsComObject(obj) Then
                Marshal.FinalReleaseComObject(obj)
            End If
        Catch
        End Try
    End Sub

    Private Async Function BuildEmbeddedOutlookItemTextAsync(
    item As Object,
    itemId As String,
    state As PstExportState,
    options As PstExportOptions,
    depth As Integer) As Task(Of String)

        Dim sb As New StringBuilder()

        AppendOutlookItemHeader(sb, item, Nothing, itemId)

        Dim bodyText As String = GetOutlookItemTextBody(item)
        sb.AppendLine("=== BODY / PRIMARY TEXT ===")
        sb.AppendLine(bodyText)

        Dim detailsText As String = BuildItemSpecificDetailsText(item)
        If Not String.IsNullOrWhiteSpace(detailsText) Then
            sb.AppendLine()
            sb.AppendLine("=== STRUCTURED DETAILS ===")
            sb.AppendLine(detailsText.Trim())
        End If

        Dim attachments As Attachments = TryGetAttachments(item)
        Dim attachmentCount As Integer = If(attachments Is Nothing, 0, attachments.Count)

        sb.AppendLine()
        sb.AppendLine("=== NESTED ATTACHMENTS ===")
        sb.AppendLine($"Count: {attachmentCount}")

        If depth < PSTExport_MaxEmbeddedDepth AndAlso attachmentCount > 0 Then
            For i As Integer = 1 To attachmentCount
                If ProgressBarModule.CancelOperation Then
                    state.Cancelled = True
                    Exit For
                End If

                Dim nestedAttachment As Attachment = Nothing

                Try
                    nestedAttachment = attachments(i)
                    sb.AppendLine()
                    sb.AppendLine(New String("-"c, 80))
                    sb.AppendLine($"NESTED ATTACHMENT {i:000}")
                    sb.AppendLine(New String("-"c, 80))
                    sb.AppendLine(Await RenderAttachmentTextAsync(
                        nestedAttachment,
                        $"{itemId}-ATT{i:000}",
                        state,
                        options,
                        depth + 1))
                Catch ex As System.Exception
                    state.PlaceholderCount += 1
                    AppendError(state, $"EMBEDDED-ATTACHMENT|ItemId={itemId}|NestedIndex={i}|Message={ex.Message}")
                    sb.AppendLine(BuildAttachmentPlaceholderText(
                        $"Nested attachment {i:000}",
                        "Embedded Outlook attachment",
                        ex.Message))
                Finally
                    ReleaseComObjectIfNeeded(nestedAttachment)
                End Try
            Next
        ElseIf attachmentCount > 0 Then
            state.PlaceholderCount += 1
            sb.AppendLine("Nested attachments were not processed because the maximum recursion depth was reached.")
        End If

        Return sb.ToString().TrimEnd()
    End Function

    Private Function CountExportableItems(folder As MAPIFolder, recurseFolders As Boolean) As Integer
        If folder Is Nothing Then Return 0

        Dim count As Integer = 0
        Dim items As Items = Nothing

        Try
            items = folder.Items
        Catch
            items = Nothing
        End Try

        If items IsNot Nothing Then
            For i As Integer = 1 To items.Count
                Dim itemObj As Object = Nothing
                Try
                    itemObj = items(i)
                    If IsExportableOutlookItem(itemObj) Then
                        count += 1
                    End If
                Catch
                Finally
                    ReleaseComObjectIfNeeded(itemObj)
                End Try
            Next
        End If

        If recurseFolders Then
            For Each subFolder As MAPIFolder In GetSortedSubFolders(folder)
                count += CountExportableItems(subFolder, True)
            Next
        End If

        Return count
    End Function

    Private Function GetSortedItemSnapshots(folder As MAPIFolder) As List(Of PstItemSnapshot)
        Dim result As New List(Of PstItemSnapshot)()

        If folder Is Nothing Then Return result

        Dim items As Items = Nothing
        Dim storeId As String = ""

        Try
            storeId = folder.Store.StoreID
        Catch
        End Try

        Try
            items = folder.Items
        Catch
            items = Nothing
        End Try

        If items Is Nothing Then Return result

        For i As Integer = 1 To items.Count
            Dim itemObj As Object = Nothing

            Try
                itemObj = items(i)
                If Not IsExportableOutlookItem(itemObj) Then
                    Continue For
                End If

                Dim entryId As String = TryGetStringProperty(itemObj, "EntryID")
                If String.IsNullOrWhiteSpace(entryId) Then
                    Continue For
                End If

                result.Add(New PstItemSnapshot() With {
                    .EntryID = entryId,
                    .StoreID = storeId,
                    .SortKey = BuildItemSortKey(itemObj, i),
                    .SubjectPreview = SafeSingleLine(GetOutlookItemSubject(itemObj), 100)
                })
            Catch
            Finally
                ReleaseComObjectIfNeeded(itemObj)
            End Try
        Next

        result.Sort(Function(a, b) String.Compare(a.SortKey, b.SortKey, StringComparison.OrdinalIgnoreCase))
        Return result
    End Function

    Private Function GetSortedSubFolders(folder As MAPIFolder) As List(Of MAPIFolder)
        Dim result As New List(Of MAPIFolder)()
        If folder Is Nothing Then Return result

        Dim folders As Folders = Nothing
        Try
            folders = folder.Folders
        Catch
            folders = Nothing
        End Try

        If folders Is Nothing Then Return result

        For Each subFolderObj As Object In folders
            Dim subFolder As MAPIFolder = TryCast(subFolderObj, MAPIFolder)
            If subFolder IsNot Nothing Then
                result.Add(subFolder)
            End If
        Next

        result.Sort(Function(a, b) String.Compare(
            If(a?.Name, ""),
            If(b?.Name, ""),
            StringComparison.OrdinalIgnoreCase))

        Return result
    End Function

    Private Function BuildItemSortKey(item As Object, fallbackIndex As Integer) As String
        Dim dt As Nullable(Of Date) = TryGetOutlookItemPrimaryDate(item)
        Dim dtKey As String = If(dt.HasValue,
            dt.Value.ToString("yyyyMMddHHmmss", Globalization.CultureInfo.InvariantCulture),
            "99999999999999")

        Dim subjectKey As String = GetOutlookItemSubject(item).Trim().ToLowerInvariant()
        Dim classKey As String = GetOutlookItemDisplayTypeName(item)
        Return dtKey & "|" & classKey & "|" & subjectKey & "|" & fallbackIndex.ToString("000000")
    End Function

    Private Function IsExportableOutlookItem(item As Object) As Boolean
        If item Is Nothing Then Return False

        Return TypeOf item Is MailItem OrElse
               TypeOf item Is AppointmentItem OrElse
               TypeOf item Is MeetingItem OrElse
               TypeOf item Is PostItem OrElse
               TypeOf item Is ReportItem OrElse
               TypeOf item Is TaskItem OrElse
               TypeOf item Is TaskRequestItem OrElse
               TypeOf item Is TaskRequestAcceptItem OrElse
               TypeOf item Is TaskRequestDeclineItem OrElse
               TypeOf item Is TaskRequestUpdateItem OrElse
               TypeOf item Is ContactItem OrElse
               TypeOf item Is DistListItem OrElse
               TypeOf item Is JournalItem OrElse
               TypeOf item Is NoteItem OrElse
               TypeOf item Is DocumentItem OrElse
               TypeOf item Is SharingItem
    End Function

    Private Sub AppendOutlookItemHeader(
        sb As StringBuilder,
        item As Object,
        owningFolder As MAPIFolder,
        itemId As String)

        sb.AppendLine($"Item ID: {itemId}")
        sb.AppendLine($"Item Type: {GetOutlookItemDisplayTypeName(item)}")
        sb.AppendLine($"Message Class: {TryGetOutlookItemMessageClass(item)}")
        sb.AppendLine($"Subject: {GetOutlookItemSubject(item)}")
        sb.AppendLine($"Participants: {GetOutlookItemParticipants(item)}")
        sb.AppendLine($"Primary Date: {FormatNullableDate(TryGetOutlookItemPrimaryDate(item))}")
        sb.AppendLine($"Outlook Folder: {If(owningFolder Is Nothing, "", TryGetFolderPath(owningFolder))}")
        sb.AppendLine()
    End Sub

    Private Function GetOutlookItemDisplayTypeName(item As Object) As String
        If item Is Nothing Then Return "Unknown"
        If TypeOf item Is MailItem Then Return "MailItem"
        If TypeOf item Is AppointmentItem Then Return "AppointmentItem"
        If TypeOf item Is MeetingItem Then Return "MeetingItem"
        If TypeOf item Is PostItem Then Return "PostItem"
        If TypeOf item Is ReportItem Then Return "ReportItem"
        If TypeOf item Is TaskItem Then Return "TaskItem"
        If TypeOf item Is TaskRequestItem Then Return "TaskRequestItem"
        If TypeOf item Is TaskRequestAcceptItem Then Return "TaskRequestAcceptItem"
        If TypeOf item Is TaskRequestDeclineItem Then Return "TaskRequestDeclineItem"
        If TypeOf item Is TaskRequestUpdateItem Then Return "TaskRequestUpdateItem"
        If TypeOf item Is ContactItem Then Return "ContactItem"
        If TypeOf item Is DistListItem Then Return "DistListItem"
        If TypeOf item Is JournalItem Then Return "JournalItem"
        If TypeOf item Is NoteItem Then Return "NoteItem"
        If TypeOf item Is DocumentItem Then Return "DocumentItem"
        If TypeOf item Is SharingItem Then Return "SharingItem"
        Return item.GetType().Name
    End Function

    Private Function GetOutlookItemSubject(item As Object) As String
        If TypeOf item Is ContactItem Then
            Dim fullName As String = TryGetStringProperty(item, "FullName")
            If Not String.IsNullOrWhiteSpace(fullName) Then Return fullName
        End If

        Dim subject As String = TryGetStringProperty(item, "Subject")
        If Not String.IsNullOrWhiteSpace(subject) Then Return subject

        If TypeOf item Is NoteItem Then
            Dim body As String = TryGetStringProperty(item, "Body")
            body = SafeSingleLine(body, 80)
            If Not String.IsNullOrWhiteSpace(body) Then Return body
        End If

        Return ""
    End Function

    Private Function GetOutlookItemParticipants(item As Object) As String
        If item Is Nothing Then Return ""

        If TypeOf item Is MailItem OrElse TypeOf item Is MeetingItem OrElse TypeOf item Is ReportItem Then
            Dim sender As String = TryGetStringProperty(item, "SenderName")
            Dim senderEmail As String = TryGetStringProperty(item, "SenderEmailAddress")
            Dim toText As String = TryGetStringProperty(item, "To")
            Dim ccText As String = TryGetStringProperty(item, "CC")

            Dim fromText As String = sender
            If Not String.IsNullOrWhiteSpace(senderEmail) Then
                If String.IsNullOrWhiteSpace(fromText) Then
                    fromText = senderEmail
                Else
                    fromText &= " <" & senderEmail & ">"
                End If
            End If

            Return $"From: {fromText}; To: {toText}; CC: {ccText}".Trim()
        End If

        If TypeOf item Is AppointmentItem Then
            Dim organizer As String = TryGetStringProperty(item, "Organizer")
            Dim req As String = TryGetStringProperty(item, "RequiredAttendees")
            Dim opt As String = TryGetStringProperty(item, "OptionalAttendees")
            Return $"Organizer: {organizer}; Required: {req}; Optional: {opt}".Trim()
        End If

        If TypeOf item Is ContactItem Then
            Return $"Emails: {TryGetStringProperty(item, "Email1Address")}; {TryGetStringProperty(item, "Email2Address")}; {TryGetStringProperty(item, "Email3Address")}".Trim()
        End If

        Return ""
    End Function

    Private Function TryGetOutlookItemPrimaryDate(item As Object) As Nullable(Of Date)
        Dim dt As Nullable(Of Date)

        dt = TryGetDateProperty(item, "ReceivedTime")
        If dt.HasValue Then Return dt

        dt = TryGetDateProperty(item, "SentOn")
        If dt.HasValue Then Return dt

        dt = TryGetDateProperty(item, "Start")
        If dt.HasValue Then Return dt

        dt = TryGetDateProperty(item, "DueDate")
        If dt.HasValue Then Return dt

        dt = TryGetDateProperty(item, "CreationTime")
        If dt.HasValue Then Return dt

        dt = TryGetDateProperty(item, "LastModificationTime")
        If dt.HasValue Then Return dt

        Return Nothing
    End Function

    Private Function TryGetOutlookItemMessageClass(item As Object) As String
        Return TryGetStringProperty(item, "MessageClass")
    End Function

    Private Function GetOutlookItemTextBody(item As Object) As String
        If item Is Nothing Then Return String.Empty

        Dim plainText As String = ""

        If TypeOf item Is MailItem Then
            Try
                plainText = NormalizeText(GetMailBody(DirectCast(item, MailItem)))
            Catch
            End Try
        Else
            plainText = NormalizeText(TryGetStringProperty(item, "Body"))
        End If

        If Not String.IsNullOrWhiteSpace(plainText) Then
            Return plainText
        End If

        Dim htmlText As String = TryGetStringProperty(item, "HTMLBody")
        If Not String.IsNullOrWhiteSpace(htmlText) Then
            Return HtmlToPlainText(htmlText)
        End If

        Return String.Empty
    End Function

    Private Function BuildItemSpecificDetailsText(item As Object) As String
        Dim sb As New StringBuilder()

        If TypeOf item Is AppointmentItem Then
            AppendDetail(sb, "Organizer", TryGetStringProperty(item, "Organizer"))
            AppendDetail(sb, "Location", TryGetStringProperty(item, "Location"))
            AppendDetail(sb, "RequiredAttendees", TryGetStringProperty(item, "RequiredAttendees"))
            AppendDetail(sb, "OptionalAttendees", TryGetStringProperty(item, "OptionalAttendees"))
            AppendDetail(sb, "Resources", TryGetStringProperty(item, "Resources"))
            AppendDetail(sb, "Start", FormatNullableDate(TryGetDateProperty(item, "Start")))
            AppendDetail(sb, "End", FormatNullableDate(TryGetDateProperty(item, "End")))
            AppendDetail(sb, "AllDayEvent", TryGetStringProperty(item, "AllDayEvent"))
            AppendDetail(sb, "BusyStatus", TryGetStringProperty(item, "BusyStatus"))
            AppendDetail(sb, "Categories", TryGetStringProperty(item, "Categories"))
        ElseIf TypeOf item Is MeetingItem Then
            AppendDetail(sb, "SenderName", TryGetStringProperty(item, "SenderName"))
            AppendDetail(sb, "SenderEmailAddress", TryGetStringProperty(item, "SenderEmailAddress"))
            AppendDetail(sb, "To", TryGetStringProperty(item, "To"))
            AppendDetail(sb, "CC", TryGetStringProperty(item, "CC"))
            AppendDetail(sb, "SentOn", FormatNullableDate(TryGetDateProperty(item, "SentOn")))
            AppendDetail(sb, "ReceivedTime", FormatNullableDate(TryGetDateProperty(item, "ReceivedTime")))
        ElseIf TypeOf item Is MailItem Then
            AppendDetail(sb, "From", GetOutlookItemParticipants(item))
            AppendDetail(sb, "ConversationTopic", TryGetStringProperty(item, "ConversationTopic"))
            AppendDetail(sb, "Importance", TryGetStringProperty(item, "Importance"))
            AppendDetail(sb, "Categories", TryGetStringProperty(item, "Categories"))
            AppendDetail(sb, "BCC", TryGetStringProperty(item, "BCC"))
        ElseIf TypeOf item Is PostItem Then
            AppendDetail(sb, "SenderName", TryGetStringProperty(item, "SenderName"))
            AppendDetail(sb, "SentOn", FormatNullableDate(TryGetDateProperty(item, "SentOn")))
            AppendDetail(sb, "Categories", TryGetStringProperty(item, "Categories"))
        ElseIf TypeOf item Is ReportItem Then
            AppendDetail(sb, "SenderName", TryGetStringProperty(item, "SenderName"))
            AppendDetail(sb, "ReceivedTime", FormatNullableDate(TryGetDateProperty(item, "ReceivedTime")))
        ElseIf TypeOf item Is TaskItem Then
            AppendDetail(sb, "Owner", TryGetStringProperty(item, "Owner"))
            AppendDetail(sb, "Status", TryGetStringProperty(item, "Status"))
            AppendDetail(sb, "StartDate", FormatNullableDate(TryGetDateProperty(item, "StartDate")))
            AppendDetail(sb, "DueDate", FormatNullableDate(TryGetDateProperty(item, "DueDate")))
            AppendDetail(sb, "PercentComplete", TryGetStringProperty(item, "PercentComplete"))
            AppendDetail(sb, "Companies", TryGetStringProperty(item, "Companies"))
            AppendDetail(sb, "Categories", TryGetStringProperty(item, "Categories"))
        ElseIf TypeOf item Is ContactItem Then
            AppendDetail(sb, "FullName", TryGetStringProperty(item, "FullName"))
            AppendDetail(sb, "CompanyName", TryGetStringProperty(item, "CompanyName"))
            AppendDetail(sb, "JobTitle", TryGetStringProperty(item, "JobTitle"))
            AppendDetail(sb, "Department", TryGetStringProperty(item, "Department"))
            AppendDetail(sb, "Email1Address", TryGetStringProperty(item, "Email1Address"))
            AppendDetail(sb, "Email2Address", TryGetStringProperty(item, "Email2Address"))
            AppendDetail(sb, "Email3Address", TryGetStringProperty(item, "Email3Address"))
            AppendDetail(sb, "BusinessTelephoneNumber", TryGetStringProperty(item, "BusinessTelephoneNumber"))
            AppendDetail(sb, "HomeTelephoneNumber", TryGetStringProperty(item, "HomeTelephoneNumber"))
            AppendDetail(sb, "MobileTelephoneNumber", TryGetStringProperty(item, "MobileTelephoneNumber"))
            AppendDetail(sb, "BusinessAddress", TryGetStringProperty(item, "BusinessAddress"))
            AppendDetail(sb, "HomeAddress", TryGetStringProperty(item, "HomeAddress"))
        ElseIf TypeOf item Is DistListItem Then
            AppendDetail(sb, "DLName", TryGetStringProperty(item, "DLName"))
            AppendDetail(sb, "MemberCount", TryGetStringProperty(item, "MemberCount"))
        ElseIf TypeOf item Is JournalItem Then
            AppendDetail(sb, "Type", TryGetStringProperty(item, "Type"))
            AppendDetail(sb, "Start", FormatNullableDate(TryGetDateProperty(item, "Start")))
            AppendDetail(sb, "Duration", TryGetStringProperty(item, "Duration"))
            AppendDetail(sb, "Companies", TryGetStringProperty(item, "Companies"))
        ElseIf TypeOf item Is NoteItem Then
            AppendDetail(sb, "Size", TryGetStringProperty(item, "Size"))
        ElseIf TypeOf item Is SharingItem Then
            AppendDetail(sb, "RemoteName", TryGetStringProperty(item, "RemoteName"))
            AppendDetail(sb, "RemotePath", TryGetStringProperty(item, "RemotePath"))
        End If

        Return sb.ToString().TrimEnd()
    End Function

    Private Sub AppendDetail(sb As StringBuilder, name As String, value As String)
        If sb Is Nothing Then Return
        If String.IsNullOrWhiteSpace(value) Then Return
        sb.AppendLine(name & ": " & value)
    End Sub

    Private Function TryGetAttachments(item As Object) As Attachments
        Try
            Return TryCast(CallByName(item, "Attachments", CallType.Get), Attachments)
        Catch
            Return Nothing
        End Try
    End Function

    Private Function TryGetStringProperty(item As Object, propertyName As String) As String
        If item Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return ""

        Try
            Dim value As Object = CallByName(item, propertyName, CallType.Get)
            If value Is Nothing Then Return ""
            Return CStr(value)
        Catch
            Return ""
        End Try
    End Function

    Private Function TryGetDateProperty(item As Object, propertyName As String) As Nullable(Of Date)
        If item Is Nothing OrElse String.IsNullOrWhiteSpace(propertyName) Then Return Nothing

        Try
            Dim value As Object = CallByName(item, propertyName, CallType.Get)
            If value Is Nothing Then Return Nothing

            Dim dt As Date
            If TypeOf value Is Date Then
                dt = DirectCast(value, Date)
                If dt <> Date.MinValue Then Return dt
            ElseIf Date.TryParse(CStr(value), dt) Then
                If dt <> Date.MinValue Then Return dt
            End If
        Catch
        End Try

        Return Nothing
    End Function

    Private Function HtmlToPlainText(html As String) As String
        If String.IsNullOrWhiteSpace(html) Then Return String.Empty

        Try
            Dim htmlDoc As New HtmlAgilityPack.HtmlDocument()
            htmlDoc.LoadHtml(html)
            Dim text As String = System.Net.WebUtility.HtmlDecode(htmlDoc.DocumentNode.InnerText)
            Return NormalizeText(text)
        Catch
            Try
                Dim stripped As String = System.Text.RegularExpressions.Regex.Replace(html, "<[^>]+>", " ")
                stripped = System.Net.WebUtility.HtmlDecode(stripped)
                Return NormalizeText(stripped)
            Catch
                Return NormalizeText(html)
            End Try
        End Try
    End Function

    Private Function FormatNullableDate(value As Nullable(Of Date)) As String
        If Not value.HasValue Then Return ""
        Return value.Value.ToString("yyyy-MM-dd HH:mm:ss", Globalization.CultureInfo.InvariantCulture)
    End Function

End Class
