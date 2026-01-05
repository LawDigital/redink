' Part of "Red Ink for Word"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.

' =============================================================================
' File: DragDropForm.vb
' Purpose: Provides a drag-and-drop interface for file selection with browse fallback.
'          Stores the selected file path and returns DialogResult.OK upon selection.
'
' Architecture:
'  - Drag-and-Drop Support: Enables AllowDrop and handles DragEnter/DragDrop events
'    to accept file drops (takes first file from drop operation).
'  - Browse Button: Opens OpenFileDialog with configurable filter (uses global
'    settings from Globals.ThisAddIn.DragDropFormFilter or default supported extensions).
'  - Customization: Form title and label text can be configured via Globals.ThisAddIn
'    properties (DragDropFormLabel).
'  - Result: Exposes SelectedFilePath property containing the chosen file path;
'    sets DialogResult.OK and closes form upon successful selection.
' =============================================================================


' Usage Examples:

' File only (default, backward compatible)
'Dim form1 As New DragDropForm()

' Directory only
'Dim form2 As New DragDropForm(DragDropMode.DirectoryOnly)

' Both file and directory
'Dim form3 As New DragDropForm(DragDropMode.FileOrDirectory)
'If form3.ShowDialog() = DialogResult.OK Then
'If form3.IsDirectory Then
' Handle directory
'Else
' Handle file
'End If
'End If

Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO

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
                Me.Text = "Drag & Drop Your File or Click Browse"
            Case DragDropMode.DirectoryOnly
                Me.Text = "Drag & Drop Your Folder or Click Browse"
            Case DragDropMode.FileOrDirectory
                Me.Text = "Drag & Drop Your File or Folder, or Click Browse"
        End Select

        If Globals.ThisAddIn.DragDropFormLabel <> "" Then
            Me.Label2.Text = Globals.ThisAddIn.DragDropFormLabel
        End If
    End Sub

    ''' <summary>
    ''' Sets the form icon from application resources on load.
    ''' </summary>
    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(My.Resources.Red_Ink_Logo)
        Dim icon As Icon = Icon.FromHandle(bmp.GetHicon())
        Me.Icon = icon
        ' Dispose bitmap to release GDI resources
        bmp.Dispose()
    End Sub

    ''' <summary>
    ''' Handles drag-enter event to accept file or directory drops with copy effect.
    ''' </summary>
    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        ' Check if the data being dragged is a file or folder
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    ''' <summary>
    ''' Handles drag-drop event to capture the first dropped file or directory and close the form with DialogResult.OK.
    ''' </summary>
    Private Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Try
            ' Retrieve the file/folder list
            Dim paths As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
            If paths IsNot Nothing AndAlso paths.Length > 0 Then
                Dim droppedPath As String = paths(0) ' Take first item
                Dim isDir As Boolean = Directory.Exists(droppedPath)
                Dim isFile As Boolean = File.Exists(droppedPath)

                Select Case _selectionMode
                    Case DragDropMode.FileOnly
                        If Not isFile Then
                            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Please drop a file, not a folder.")
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
            End If
        Catch ex As System.Exception
            MessageBox.Show($"Error: {ex.Message}")
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

                ' Default filter covering text, document, and presentation formats
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.pptx;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Text Files (*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm)|*.txt;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm|" &
                             "Rich Text Files (*.rtf)|*.rtf|" &
                             "Word Documents (*.doc;*.docx)|*.doc;*.docx|" &
                             "PDF Files (*.pdf)|*.pdf|" &
                             "Powerpoint Files (*.pptx)|*.pptx|" &
                             "All Files (*.*)|*.*"

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