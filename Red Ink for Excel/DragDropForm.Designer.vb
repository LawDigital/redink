' Part of "Red Ink for Excel"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'

Imports System.Drawing

''' <summary>
''' Designer-generated partial class providing the visual layout for the drag-and-drop dialog.
''' </summary>
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DragDropForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    ''' <summary>
    ''' Releases managed resources when disposing the designer components container.
    ''' </summary>
    ''' <param name="disposing">True when invoked from managed code; False when finalizing.</param>
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
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(12, 166)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(84, 35)
        Me.btnBrowse.TabIndex = 0
        Me.btnBrowse.Text = "Browse ..."
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(12, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(198, 20)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Drag and drop here ..."
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(12, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(320, 80)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Supported are Text Files (*.txt; *.ini; *.csv; *.log; *.json; *.xml; *.html; *.ht" &
    "m), RTF Files (*.rtf), Word Documents (*.doc; *.docx) and PDF Files (*.pdf)"
        '
        'DragDropForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(354, 224)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnBrowse)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DragDropForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Load external file"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnBrowse As Windows.Forms.Button

    Friend WithEvents Label1 As Windows.Forms.Label

    Friend WithEvents Label2 As Windows.Forms.Label

End Class