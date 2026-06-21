' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: ThisAddIn.FileHelpers.vb
' Purpose: Provides file path acquisition via drag-and-drop form and content
'          loading utilities for supported document types with optional OCR.
'
' Architecture:
' - This file contributes to the partial class ThisAddIn.
' - Relies on external helper methods (e.g., RemoveCR, ShowCustomMessageBox,
'   ExpandEnvironmentVariables, ReadTextFile, ReadRtfAsText, ReadWordDocument,
'   ReadPdfAsText, ReadBinaryFileViaLLM) from SharedLibrary.SharedLibrary.SharedMethods.
' - Uses a modal DragDropForm UI to capture a user-selected file path.
' - Validates file existence before returning or processing.
' - Asynchronous content loading for PDF to allow OCR and user interaction flags.
' - File type dispatch via Select Case on extension for text extraction.
' - Binary/media files (images, audio, video) are sent to the LLM as binary objects.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Diagnostics
Imports System.IO
Imports System.Threading.Tasks
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedMethods

Partial Public Class ThisAddIn

    ''' <summary>
    ''' Retrieves content from a web URL and stores it as a temporary file.
    ''' Documents are downloaded as original files; websites are retrieved via the shared WebView2 importer and stored as text.
    ''' </summary>
    Public Async Function CreateTempFileFromUrlAsync(url As String) As Task(Of String)
        Return Await SharedLibrary.SharedLibrary.SharedMethods.CreateTempFileFromUrlAsync(
            url,
            SharedLibrary.SharedLibrary.SharedMethods.GetDefaultWebImportSupportedExtensions(INI_AllowLegacyDocFiles, False))
    End Function

    ''' <summary>
    ''' Opens a drag-and-drop selection form, obtains a file path, normalizes it,
    ''' validates existence, and returns the absolute path or empty string on failure.
    ''' </summary>
    ''' <returns>Full file path if found; otherwise empty string.</returns>
    Public Function GetFileName() As String
        Dim filePath As String = ""
        Try
            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        ' User cancelled or closed form
                        Return String.Empty
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = Path.GetFullPath(filePath)
            If Not File.Exists(filePath) Then
                ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return ""
            End If
            Return filePath

        Catch ex As System.Exception
            ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' Retrieves textual content from a supported file, returning extended result with PDF completeness info.
    ''' Supports optional worksheet selection for Excel workbooks.
    ''' </summary>
    ''' <param name="optionalFilePath">File path to load; environment variables are expanded when provided.</param>
    ''' <param name="Silent">Suppresses UI error/notification messages when set to True.</param>
    ''' <param name="DoOCR">Enables OCR while reading PDF files when True.</param>
    ''' <param name="AskUser">Indicates whether PDF processing may prompt the user.</param>
    ''' <param name="AskWorksheetSelection">
    ''' For Excel workbooks, when <c>True</c> and <paramref name="Silent"/> is <c>False</c>,
    ''' allows the user to choose one worksheet or all worksheets.
    ''' </param>
    ''' <returns>
    ''' A <see cref="FileReadResult"/> containing the extracted content and whether a PDF may be incomplete.
    ''' </returns>
    Public Async Function GetFileContentEx(Optional ByVal optionalFilePath As String = Nothing,
                                            Optional Silent As Boolean = False,
                                            Optional DoOCR As Boolean = False,
                                            Optional AskUser As Boolean = True,
                                            Optional AskWorksheetSelection As Boolean = False) As Task(Of FileReadResult)
        Dim result As New FileReadResult()
        Dim filePath As String = ""

        Try
            If optionalFilePath IsNot Nothing Then
                filePath = ExpandEnvironmentVariables(optionalFilePath)
            End If

            If String.IsNullOrWhiteSpace(filePath) Then
                Using form As New DragDropForm()
                    If form.ShowDialog() = DialogResult.OK Then
                        filePath = form.SelectedFilePath
                    Else
                        Return result
                    End If
                End Using
            End If

            filePath = RemoveCR(filePath.Trim())
            filePath = System.IO.Path.GetFullPath(filePath)

            If Not File.Exists(filePath) Then
                If Not Silent Then ShowCustomMessageBox($"The file '{filePath}' was not found.")
                Return result
            End If

            If Not String.IsNullOrWhiteSpace(filePath) AndAlso IO.File.Exists(filePath) Then
                Dim ext As String = IO.Path.GetExtension(filePath).ToLowerInvariant()
                Dim FromFile As String = ""

                Select Case ext
                    Case ".txt", ".ini", ".csv", ".log", ".json", ".xml", ".html", ".htm",
                         ".md", ".yaml", ".yml",
                         ".vb", ".cs", ".js", ".ts", ".py", ".java", ".cpp", ".c", ".h", ".sql"
                        FromFile = ReadTextFile(filePath)
                    Case ".rtf"
                        FromFile = ReadRtfAsText(filePath)
                    Case ".doc"
                        If INI_AllowLegacyDocFiles Then
                            FromFile = ReadWordDocument(filePath)
                        Else
                            FromFile = "Error: File type not supported (disabled for security)."
                        End If
                    Case ".docx"
                        FromFile = ReadDocxSandboxed(filePath)
                    Case ".xlsx"
                        FromFile = ReadXlsxSandboxed(filePath, Silent, AskWorksheetSelection)
                        If String.Equals(FromFile, XlsxSelectionCancelledMarker, StringComparison.Ordinal) Then
                            result.UserCancelled = True
                            result.Content = ""
                            Return result
                        End If
                    Case ".pptx"
                        FromFile = ReadPptxSandboxed(filePath)
                    Case ".pdf"
                        Dim pdfResult = Await ReadPdfAsTextEx(filePath, True, DoOCR, AskUser, _context)
                        FromFile = pdfResult.Content
                        result.PdfMayBeIncomplete = pdfResult.OcrWasSkippedDueToHeuristics
                    Case ".eml"
                        FromFile = ReadEmlSandboxed(filePath)
                    Case ".msg"
                        FromFile = ReadMsgSandboxed(filePath)
                    Case Else
                        ' Check if this is a binary/media file the model can handle directly
                        If IsBinaryMediaExtension(ext) Then
                            Dim taskFlag = TaskFlagForExtension(ext)
                            If IsBinaryMediaSupported(_context, ext, taskFlag) Then
                                Try
                                    FromFile = Await ReadBinaryFileViaLLM(filePath, _context, "", AskUser, taskFlag)
                                    If String.IsNullOrWhiteSpace(FromFile) Then
                                        FromFile = ""
                                    End If
                                Catch ex As System.Exception
                                    FromFile = ""
                                    Debug.WriteLine("Binary media extraction failed for '" & filePath & "': " & ex.Message)
                                End Try
                            Else
                                FromFile = "Error: The file type '" & ext & "' is not supported by your current model configuration."
                            End If
                        Else
                            FromFile = "Error: File type not supported."
                        End If
                End Select

                If FromFile.StartsWith("Error") AndAlso Len(FromFile) < 100 AndAlso Not Silent Then
                    ShowCustomMessageBox(FromFile)
                    result.Content = ""
                Else
                    result.Content = FromFile
                End If
            End If

        Catch ex As System.Exception
            If Not Silent Then ShowCustomMessageBox($"An error occurred reading the file '{filePath}': {ex.Message}")
            result.Content = ""
        End Try

        Return result
    End Function

    ''' <summary>
    ''' Retrieves textual content from a supported file (backward compatible wrapper).
    ''' Supports optional worksheet selection for Excel workbooks.
    ''' </summary>
    Public Async Function GetFileContent(Optional ByVal optionalFilePath As String = Nothing,
                                         Optional Silent As Boolean = False,
                                         Optional DoOCR As Boolean = False,
                                         Optional AskUser As Boolean = True,
                                         Optional AskWorksheetSelection As Boolean = False) As Task(Of String)
        Dim result = Await GetFileContentEx(optionalFilePath, Silent, DoOCR, AskUser, AskWorksheetSelection)
        Return result.Content
    End Function

End Class
