Imports System.Windows.Forms

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmAIChat
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.lblInstructions = New System.Windows.Forms.Label()
        Me.txtChatHistory = New MyTextBox()
        Me.txtUserInput = New MyTextBox()
        Me.SuspendLayout()
        '
        'lblInstructions
        '
        Me.lblInstructions.AutoSize = True
        Me.lblInstructions.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInstructions.Location = New System.Drawing.Point(15, 13)
        Me.lblInstructions.Name = "lblInstructions"
        Me.lblInstructions.Size = New System.Drawing.Size(90, 20)
        Me.lblInstructions.TabIndex = 0
        Me.lblInstructions.Text = "Your AI Chat"
        '
        'txtChatHistory
        '
        Me.txtChatHistory.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChatHistory.Location = New System.Drawing.Point(19, 47)
        Me.txtChatHistory.Multiline = True
        Me.txtChatHistory.Name = "txtChatHistory"
        Me.txtChatHistory.ReadOnly = True
        Me.txtChatHistory.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtChatHistory.Size = New System.Drawing.Size(611, 245)
        Me.txtChatHistory.TabIndex = 1
        '
        'txtUserInput
        '
        Me.txtUserInput.Font = New System.Drawing.Font("Segoe UI", 10.2!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUserInput.Location = New System.Drawing.Point(19, 298)
        Me.txtUserInput.Multiline = True
        Me.txtUserInput.Name = "txtUserInput"
        Me.txtUserInput.Size = New System.Drawing.Size(611, 63)
        Me.txtUserInput.TabIndex = 2
        '
        'frmAIChat
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(649, 474)
        Me.Controls.Add(Me.txtUserInput)
        Me.Controls.Add(Me.txtChatHistory)
        Me.Controls.Add(Me.lblInstructions)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.KeyPreview = True
        Me.Name = "frmAIChat"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lblInstructions As System.Windows.Forms.Label
    Friend WithEvents txtChatHistory As System.Windows.Forms.TextBox
    Friend WithEvents txtUserInput As System.Windows.Forms.TextBox

End Class

Public Class MyTextBox
    Inherits TextBox

    Protected Overrides Sub OnMouseWheel(e As MouseEventArgs)
        ' Only change the font size if the Ctrl key is pressed.
        If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            Dim currentSize As Single = Me.Font.Size
            Dim newSize As Single = currentSize

            ' Increase or decrease font size based on wheel direction.
            If e.Delta > 0 Then
                newSize += 1
            ElseIf e.Delta < 0 Then
                newSize -= 1
            End If

            ' Prevent the font size from becoming too small.
            If newSize < 2 Then newSize = 2

            Try
                ' Update the font while preserving the font family and style.
                Me.Font = New System.Drawing.Font(Me.Font.FontFamily, newSize, Me.Font.Style)
            Catch ex As System.Exception
                Global.SharedLibrary.SharedLibrary.SharedMethods.ShowCustomMessageBox("Error changing font size: " & ex.Message)
            End Try
        End If

        ' Always call the base method for standard processing.
        MyBase.OnMouseWheel(e)
    End Sub
End Class

