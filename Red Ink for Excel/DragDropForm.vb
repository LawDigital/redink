' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DragDropForm.vb
' Purpose: Provides a drag-and-drop interface for file or directory selection with browse fallback.
'          Stores the selected path and returns DialogResult.OK upon selection.
'
' Architecture:
'  - Drag-and-Drop Support: Enables AllowDrop and handles DragEnter/DragDrop events
'    to accept file or folder drops (takes first item from drop operation).
'  - Browse Button: Opens OpenFileDialog or FolderBrowserDialog based on selection mode.
'  - Selection Modes: Supports FileOnly, DirectoryOnly, or FileOrDirectory via DragDropMode enum.
'  - Customization: Form title and label text can be configured via Globals.ThisAddIn
'    properties (DragDropFormLabel).
'  - Result: Exposes SelectedFilePath property containing the chosen path;
'    sets DialogResult.OK and closes form upon successful selection.
' =============================================================================

Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports SharedLibrary.SharedLibrary

''' <summary>
''' Specifies what type of path the DragDropForm should accept.
''' </summary>
Public Enum DragDropMode
    ''' <summary>Accept only files.</summary>
    FileOnly = 0
    ''' <summary>Accept only directories.</summary>
    DirectoryOnly = 1
    ''' <summary>Accept both files and directories.</summary>
    FileOrDirectory = 2
End Enum

Public Class DragDropForm

    Private _selectedFilePath As String = String.Empty
    Private _selectionMode As DragDropMode = DragDropMode.FileOnly
    Private _dialogOwnerScope As System.IDisposable = Nothing

    ' Layout constants
    Private Const LabelToButtonSpacing As Integer = 20
    Private Const ButtonToFormBottomSpacing As Integer = 24

    ''' <summary>
    ''' Gets the file or directory path selected by the user via drag-and-drop or browse dialog.
    ''' </summary>
    Public ReadOnly Property SelectedFilePath As String
        Get
            Return _selectedFilePath
        End Get
    End Property

    ''' <summary>
    ''' Gets whether the selected path is a directory.
    ''' </summary>
    Public ReadOnly Property IsDirectory As Boolean
        Get
            Return Directory.Exists(_selectedFilePath)
        End Get
    End Property

    ''' <summary>
    ''' Gets the current selection mode.
    ''' </summary>
    Public ReadOnly Property SelectionMode As DragDropMode
        Get
            Return _selectionMode
        End Get
    End Property

    Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
        MyBase.OnHandleCreated(e)

        If _dialogOwnerScope Is Nothing Then
            _dialogOwnerScope = SharedMethods.PushDialogOwner(Me)
        End If
    End Sub

    Protected Overrides Sub OnHandleDestroyed(e As System.EventArgs)
        Dim scope As System.IDisposable = _dialogOwnerScope
        _dialogOwnerScope = Nothing

        If scope IsNot Nothing Then
            Try
                scope.Dispose()
            Catch
            End Try
        End If

        MyBase.OnHandleDestroyed(e)
    End Sub

    ''' <summary>
    ''' Initializes the form with drag-and-drop enabled and optional custom label text.
    ''' Defaults to file-only mode.
    ''' </summary>
    Public Sub New()
        Me.New(DragDropMode.FileOnly)
    End Sub

    ''' <summary>
    ''' Initializes the form with drag-and-drop enabled, optional custom label text, and specified selection mode.
    ''' </summary>
    ''' <param name="mode">Specifies whether to accept files only, directories only, or both.</param>
    Public Sub New(mode As DragDropMode)
        InitializeComponent()
        _selectionMode = mode

        ' Ensure drag and drop is enabled
        Me.AllowDrop = True

        ' Adjust form title based on mode
        Select Case _selectionMode
            Case DragDropMode.FileOnly
                Me.Text = "Your File/Link"
            Case DragDropMode.DirectoryOnly
                Me.Text = "Your Folder"
            Case DragDropMode.FileOrDirectory
                Me.Text = "Your File/Folder/Link"
        End Select

        ' Update the supported-formats label to stay in sync with the actual file filter
        If Globals.ThisAddIn.DragDropFormLabel <> "" Then
            Me.Label2.Text = Globals.ThisAddIn.DragDropFormLabel
        Else
            Me.Label2.Text = GetDefaultSupportedFormatsText()
        End If

        ' Resize the form so the label, button, and bottom margin all fit
        AdjustFormLayout()
    End Sub

    ''' <summary>
    ''' Repositions the browse button below Label2 and resizes the form height to fit all content.
    ''' </summary>
    Private Sub AdjustFormLayout()
        ' Let the label compute its auto-sized height
        Me.Label2.PerformLayout()

        ' Position the button below the label
        Me.btnBrowse.Top = Me.Label2.Bottom + LabelToButtonSpacing

        ' Resize the form to fit the button plus bottom margin
        Me.ClientSize = New Size(Me.ClientSize.Width, Me.btnBrowse.Bottom + ButtonToFormBottomSpacing)
    End Sub

    ''' <summary>
    ''' Builds the default "Supported are ..." label text based on the current selection mode and legacy-doc setting.
    ''' This keeps the UI label in sync with the actual file filter used by BrowseForFile.
    ''' </summary>
    Private Function GetDefaultSupportedFormatsText() As String
        Select Case _selectionMode
            Case DragDropMode.DirectoryOnly
                Return "Drop or browse for a folder."

            Case Else
                ' Build the description from the same extensions used in BrowseForFile
                Dim parts As New List(Of String)

                If ThisAddIn.INI_AllowLegacyDocFiles Then
                    parts.Add("Text Files (*.txt; *.ini; *.csv; *.log; *.json; *.xml; *.html; *.htm; *.md; *.yaml; *.yml)")
                    parts.Add("RTF Files (*.rtf)")
                    parts.Add("Word Documents (*.doc; *.docx)")
                    parts.Add("Excel Workbooks (*.xlsx)")
                    parts.Add("PowerPoint Files (*.pptx)")
                    parts.Add("PDF Files (*.pdf)")
                    parts.Add("Email Files (*.msg; *.eml)")
                    parts.Add("Source Code (*.vb; *.cs; *.js; *.ts; *.py; *.java; *.cpp; *.c; *.h; *.sql)")
                Else
                    parts.Add("Text Files (*.txt; *.ini; *.csv; *.log; *.json; *.xml; *.html; *.htm; *.md; *.yaml; *.yml)")
                    parts.Add("RTF Files (*.rtf)")
                    parts.Add("Word Documents (*.docx)")
                    parts.Add("Excel Workbooks (*.xlsx)")
                    parts.Add("PowerPoint Files (*.pptx)")
                    parts.Add("PDF Files (*.pdf)")
                    parts.Add("Email Files (*.msg; *.eml)")
                    parts.Add("Source Code (*.vb; *.cs; *.js; *.ts; *.py; *.java; *.cpp; *.c; *.h; *.sql)")
                End If

                ' Join with commas and " and " before the last item
                If parts.Count <= 1 Then
                    Return "Supported are " & parts(0) & "."
                End If

                Dim allButLast As String = String.Join(", ", parts.Take(parts.Count - 1))
                Return "Supported are " & allButLast & " and " & parts.Last() & "."
        End Select
    End Function

    ''' <summary>
    ''' Sets the form icon from application resources on load.
    ''' </summary>
    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
        Dim icon As Icon = Icon.FromHandle(bmp.GetHicon())
        Me.Icon = icon
        ' Dispose bitmap to release GDI resources
        bmp.Dispose()

        ' Ensure the form appears above TopMost progress bars that may be running
        ' on a separate STA thread (e.g., during multi-trigger file loading).
        ' Must stay TopMost because the progress form on its own thread keeps
        ' reclaiming the foreground via its TopMost=True property.
        Me.TopMost = True
        Me.BringToFront()
        Me.Activate()
    End Sub

    ''' <summary>
    ''' Handles drag-enter event to accept file, directory, or internet-link drops with copy effect.
    ''' </summary>
    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data Is Nothing Then
            e.Effect = DragDropEffects.None
            Return
        End If

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
            Return
        End If

        If _selectionMode <> DragDropMode.DirectoryOnly Then
            Dim droppedUrl As String = ""
            If SharedLibrary.SharedLibrary.SharedMethods.TryGetDroppedInternetLink(e.Data, droppedUrl) Then
                e.Effect = DragDropEffects.Copy
                Return
            End If
        End If

        e.Effect = DragDropEffects.None
    End Sub

    ''' <summary>
    ''' Handles drag-drop event to capture the first dropped file, directory, or internet link.
    ''' Internet links are downloaded or retrieved to temp files.
    ''' </summary>
    Private Async Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Try
            If e.Data Is Nothing Then
                Return
            End If

            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim paths As String() = TryCast(e.Data.GetData(DataFormats.FileDrop), String())
                If paths IsNot Nothing AndAlso paths.Length > 0 Then
                    Dim droppedPath As String = paths(0)
                    Dim droppedUrl As String = ""

                    If _selectionMode <> DragDropMode.DirectoryOnly AndAlso
                       String.Equals(Path.GetExtension(droppedPath), ".url", StringComparison.OrdinalIgnoreCase) Then

                        If SharedLibrary.SharedLibrary.SharedMethods.TryReadInternetShortcutUrl(droppedPath, droppedUrl) Then
                            If Await TryHandleDroppedInternetLinkAsync(droppedUrl) Then
                                Return
                            End If
                        End If

                        SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("The dropped internet shortcut could not be read. You may want to try again.")
                        Return
                    End If

                    Dim isDir As Boolean = Directory.Exists(droppedPath)
                    Dim isFile As Boolean = File.Exists(droppedPath)

                    Select Case _selectionMode
                        Case DragDropMode.FileOnly
                            If Not isFile Then
                                SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Please drop a file/link, not a folder.")
                                Return
                            End If

                        Case DragDropMode.DirectoryOnly
                            If Not isDir Then
                                SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Please drop a folder, not a file.")
                                Return
                            End If

                        Case DragDropMode.FileOrDirectory
                            ' Accept both - no validation needed
                    End Select

                    _selectedFilePath = droppedPath
                    Me.DialogResult = DialogResult.OK
                    Me.Close()
                    Return
                End If
            End If

            If _selectionMode <> DragDropMode.DirectoryOnly Then
                Dim droppedUrl As String = ""
                If SharedLibrary.SharedLibrary.SharedMethods.TryGetDroppedInternetLink(e.Data, droppedUrl) Then
                    Await TryHandleDroppedInternetLinkAsync(droppedUrl)
                    Return
                End If
            End If

        Catch ex As System.Exception
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox($"Error: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Retrieves a dropped internet link, stores it as a temp file, and selects that file.
    ''' </summary>
    Private Async Function TryHandleDroppedInternetLinkAsync(url As String) As System.Threading.Tasks.Task(Of Boolean)
        If _selectionMode = DragDropMode.DirectoryOnly Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Please drop a folder, not an internet link.")
            Return False
        End If

        If Globals.ThisAddIn Is Nothing Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("The dropped internet link could not be processed. You may want to try again.")
            Return False
        End If

        Dim splash As SplashScreenWorks = Nothing
        Dim tempFilePath As String = ""

        Try
            splash = ShowBusySplash("Importing website or file link, please wait...")
            Await System.Threading.Tasks.Task.Yield()

            tempFilePath = Await Globals.ThisAddIn.CreateTempFileFromUrlAsync(url)

        Finally
            CloseBusySplash(splash)
        End Try

        If String.IsNullOrWhiteSpace(tempFilePath) OrElse Not File.Exists(tempFilePath) Then
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("The dropped internet link could not be retrieved. You may want to try again.")
            Return False
        End If

        _selectedFilePath = tempFilePath
        Me.DialogResult = DialogResult.OK
        Me.Close()
        Return True
    End Function

    Private Function ShowBusySplash(message As String) As SplashScreenWorks
        Dim splash As New SplashScreenWorks(message)
        splash.StartPosition = FormStartPosition.CenterScreen
        splash.ShowInTaskbar = False
        splash.TopMost = True
        splash.Show(Me)
        splash.BringToFront()
        splash.Refresh()
        Return splash
    End Function

    Private Sub CloseBusySplash(splash As SplashScreenWorks)
        If splash Is Nothing Then
            Return
        End If

        Try
            If splash.IsDisposed Then
                Return
            End If

            splash.Close()
            splash.Dispose()
        Catch
        End Try
    End Sub

    ''' <summary>
    ''' Opens file or folder browse dialog based on selection mode.
    ''' </summary>
    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Select Case _selectionMode
            Case DragDropMode.FileOnly
                BrowseForFile()

            Case DragDropMode.DirectoryOnly
                BrowseForFolder()

            Case DragDropMode.FileOrDirectory
                ' Show choice dialog for file or folder selection
                Dim result As Integer = SharedLibrary.SharedLibrary.SharedMethods.ShowCustomYesNoBox("What do you want to browse for?", "File", "Folder")
                If result = 1 Then
                    BrowseForFile()
                ElseIf result = 2 Then
                    BrowseForFolder()
                End If
        End Select
    End Sub

    ''' <summary>
    ''' Opens OpenFileDialog to select a file.
    ''' </summary>
    Private Sub BrowseForFile()
        Using ofd As New OpenFileDialog()

            If Globals.ThisAddIn.DragDropFormFilter = "" Then

                ' Default filter — legacy formats (.doc) only shown when INI_AllowLegacyDocFiles = True
                If ThisAddIn.INI_AllowLegacyDocFiles Then
                    ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|" &
                                 "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Rich Text Files (*.rtf)|*.rtf|" &
                                 "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                                 "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                 "PowerPoint Files (*.pptx)|*.pptx|" &
                                 "PDF Files (*.pdf)|*.pdf|" &
                                 "Email Files (*.msg;*.eml)|*.msg;*.eml|" &
                                 "Source Code|*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|" &
                                 "All Files (*.*)|*.*"
                Else
                    ofd.Filter = "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|" &
                                 "Text Files|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml|" &
                                 "Rich Text Files (*.rtf)|*.rtf|" &
                                 "Word Documents (*.docx)|*.docx|" &
                                 "Excel Workbooks (*.xlsx)|*.xlsx|" &
                                 "PowerPoint Files (*.pptx)|*.pptx|" &
                                 "PDF Files (*.pdf)|*.pdf|" &
                                 "Email Files (*.msg;*.eml)|*.msg;*.eml|" &
                                 "Source Code|*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|" &
                                 "All Files (*.*)|*.*"
                End If

            Else

                ofd.Filter = Globals.ThisAddIn.DragDropFormFilter

            End If

            ofd.Title = "Select a File"
            ofd.Multiselect = False

            If ofd.ShowDialog() = DialogResult.OK Then
                _selectedFilePath = ofd.FileName
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Opens FolderBrowserDialog to select a directory.
    ''' </summary>
    Private Sub BrowseForFolder()
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select a Folder"
            fbd.ShowNewFolderButton = True

            If fbd.ShowDialog() = DialogResult.OK Then
                _selectedFilePath = fbd.SelectedPath
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

End Class
