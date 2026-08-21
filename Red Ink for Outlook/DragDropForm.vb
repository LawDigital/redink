Option Explicit On
Option Strict On

Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary

Public Enum DragDropMode
    FileOnly = 0
    DirectoryOnly = 1
    FileOrDirectory = 2
End Enum

Partial Public Class DragDropForm
    Inherits System.Windows.Forms.Form

    Protected Overrides Sub OnShown(e As System.EventArgs)
        MyBase.OnShown(e)
        Me.TopMost = True
        SharedLibrary.SharedLibrary.SharedMethods.ForceDialogToForeground(Me)
        SharedLibrary.SharedLibrary.SharedMethods.AttachForeignForegroundWatchdog(Me)
    End Sub


    Private _selectedFilePath As String = String.Empty
    Private _selectionMode As DragDropMode = DragDropMode.FileOnly
    Private _dialogOwnerScope As System.IDisposable = Nothing

    Private Const LabelToButtonSpacing As Integer = 20
    Private Const ButtonToFormBottomSpacing As Integer = 24

    Public ReadOnly Property SelectedFilePath As String
        Get
            Return _selectedFilePath
        End Get
    End Property

    Public ReadOnly Property IsDirectory As Boolean
        Get
            Return Directory.Exists(_selectedFilePath)
        End Get
    End Property

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

    Public Sub New()
        Me.New(DragDropMode.FileOnly)
    End Sub

    Public Sub New(mode As DragDropMode)
        InitializeComponent()
        _selectionMode = mode

        Me.AllowDrop = True

        Select Case _selectionMode
            Case DragDropMode.FileOnly
                Me.Text = "Your File/Link"
                Me.Label2.Text = "Drop or browse for a file or link."
            Case DragDropMode.DirectoryOnly
                Me.Text = "Your Folder"
                Me.Label2.Text = "Drop or browse for a folder."
            Case DragDropMode.FileOrDirectory
                Me.Text = "Your File/Folder/Link"
                Me.Label2.Text = "Drop or browse for a file, folder, or link."
        End Select

        AdjustFormLayout()
    End Sub

    Public Sub SetInstructionText(text As String)
        Me.Label2.Text = If(text, "")
        AdjustFormLayout()
    End Sub

    Private Sub AdjustFormLayout()
        Me.Label2.PerformLayout()
        Me.btnBrowse.Top = Me.Label2.Bottom + LabelToButtonSpacing
        Me.ClientSize = New Size(Me.ClientSize.Width, Me.btnBrowse.Bottom + ButtonToFormBottomSpacing)
    End Sub

    Private Sub DragDropForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
        Dim iconValue As Icon = Icon.FromHandle(bmp.GetHicon())
        Me.Icon = iconValue
        bmp.Dispose()

        Me.TopMost = True
        Me.BringToFront()
        Me.Activate()

    End Sub

    Private Sub DragDropForm_DragEnter(sender As Object, e As DragEventArgs) Handles MyBase.DragEnter
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

    Private Async Sub DragDropForm_DragDrop(sender As Object, e As DragEventArgs) Handles MyBase.DragDrop
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

        Catch ex As Exception
            SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error: " & ex.Message)
        End Try
    End Sub

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

    Private Sub btnBrowse_Click(sender As Object, e As EventArgs) Handles btnBrowse.Click
        Select Case _selectionMode
            Case DragDropMode.FileOnly
                BrowseForFile()

            Case DragDropMode.DirectoryOnly
                BrowseForFolder()

            Case DragDropMode.FileOrDirectory
                Dim result As Integer = SharedMethods.ShowCustomYesNoBox(
                    "What do you want to browse for?",
                    "File",
                    "Folder")

                If result = 1 Then
                    BrowseForFile()
                ElseIf result = 2 Then
                    BrowseForFolder()
                End If
        End Select
    End Sub

    Private Sub BrowseForFile()
        Using ofd As New OpenFileDialog()
            If ThisAddIn.INI_AllowLegacyDocFiles Then
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.doc;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|All Files (*.*)|*.*"
            Else
                ofd.Filter = "Supported Files|*.txt;*.rtf;*.docx;*.pdf;*.xlsx;*.pptx;*.msg;*.eml;*.ini;*.csv;*.log;*.json;*.xml;*.html;*.htm;*.md;*.yaml;*.yml;*.vb;*.cs;*.js;*.ts;*.py;*.java;*.cpp;*.c;*.h;*.sql|All Files (*.*)|*.*"
            End If

            ofd.Title = "Select a File"
            ofd.Multiselect = False

            Dim __safeDialogOwner297 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner297 IsNot Nothing, ofd.ShowDialog(__safeDialogOwner297), ofd.ShowDialog()) = DialogResult.OK Then
                _selectedFilePath = ofd.FileName
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

    Private Sub BrowseForFolder()
        Using fbd As New FolderBrowserDialog()
            fbd.Description = "Select a Folder"
            fbd.ShowNewFolderButton = True

            Dim __safeDialogOwner310 As System.Windows.Forms.IWin32Window = SharedLibrary.SharedLibrary.SharedMethods.ResolveSameThreadDialogOwner()
            If If(__safeDialogOwner310 IsNot Nothing, fbd.ShowDialog(__safeDialogOwner310), fbd.ShowDialog()) = DialogResult.OK Then
                _selectedFilePath = fbd.SelectedPath
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End If
        End Using
    End Sub

End Class
