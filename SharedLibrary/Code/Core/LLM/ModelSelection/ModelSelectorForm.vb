' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: ModelSelectorForm.vb
' Purpose: Implements a modal Windows Forms dialog that lets the user select an
'          alternative model configuration loaded from an INI file, optionally
'          including a "Default" entry.
'
' Architecture:
'  - UI Layout: Uses a single-column `TableLayoutPanel`:
'      Row 1: Title label
'      Row 2: List of models (`ListBox`)
'      Row 3: Optional reset checkbox (shown only when `OptionText` is non-empty)
'      Row 4: OK/Cancel buttons in a `FlowLayoutPanel`
'  - Model Source: Calls `LoadAlternativeModels(iniFilePath, context)` to load
'    `ModelConfig` entries from an INI file.
'  - Selection Output:
'      - `UseDefault=True` when the "Default" entry is selected.
'      - `SelectedModel` set to the chosen `ModelConfig` when a non-default entry is selected.
'  - State Coupling: Updates shared flags `OptionChecked` and `originalConfigLoaded`
'    (defined in the model selection subsystem) based on user checkbox state and selection.
'  - DPI Scaling: Uses `AutoScaleMode.Dpi` and applies an additional scale factor in
'    `OnHandleCreated` based on `DeviceDpi`.
' =============================================================================

Option Strict On
Option Explicit On

Imports System.Drawing
Imports System.Windows.Forms
Imports SharedLibrary.SharedLibrary.SharedContext
Imports SharedLibrary.SharedLibrary.SharedMethods

Namespace SharedLibrary
    ''' <summary>
    ''' Modal dialog for selecting an alternative model configuration loaded from an INI file.
    ''' </summary>
    Public Class ModelSelectorForm
        Inherits Form

        ''' <summary>Title/description label displayed above the model list.</summary>
        Private lblTitle As System.Windows.Forms.Label

        ''' <summary>List box containing the selectable model entries (and optionally the "Default" entry).</summary>
        Private lstModels As ListBox

        ''' <summary>Optional checkbox controlling whether configuration should be reset to default after use.</summary>
        Private chkReset As System.Windows.Forms.CheckBox

        ''' <summary>OK button applying the selection.</summary>
        Private btnOK As Button

        ''' <summary>Cancel button closing the dialog without applying changes.</summary>
        Private btnCancel As Button

        ''' <summary>Alternative model configurations loaded from the INI file.</summary>
        Private alternativeModels As List(Of ModelConfig)

        ''' <summary>Indicates whether the ListBox contains a "Default" entry at index 0.</summary>
        Private hasDefaultEntry As Boolean

        ''' <summary>Padding used by callers for button text layout.</summary>
        Public Shared ReadOnly ButtonTextPadding As System.Windows.Forms.Padding = New System.Windows.Forms.Padding(8, 4, 8, 4)

        ''' <summary>
        ''' Gets the selected alternative model configuration when a non-default entry is chosen; otherwise <c>Nothing</c>.
        ''' </summary>
        Public Property SelectedModel As ModelConfig = Nothing

        ''' <summary>
        ''' Gets whether the "Default" configuration entry was selected.
        ''' </summary>
        Public Property UseDefault As Boolean = True

        ''' <summary>
        ''' Initializes a new instance of the model selection dialog.
        ''' </summary>
        ''' <param name="iniFilePath">Path to the INI file containing model configurations.</param>
        ''' <param name="context">Shared context used to load settings and display the current default model name.</param>
        ''' <param name="Title">Form title text.</param>
        ''' <param name="ListType">Label text shown above the list box.</param>
        ''' <param name="OptionText">Checkbox label; if empty, the checkbox is not shown.</param>
        ''' <param name="UseCase">
        ''' Selection mode flag used by callers:
        ''' 1 = includes a default entry in the list;
        ''' 2 = does not include a default entry.
        ''' </param>
        Public Sub New(ByVal iniFilePath As String, ByVal context As ISharedContext, ByVal Title As String, ByVal ListType As String, ByVal OptionText As String, Optional UseCase As Integer = 1)

            ' UseCase 1 = Model Selection (with Default) UseCase 2 = Model Selection (without Default)

            OptionChecked = True

            ' --- Enable DPI and font scaling ---
            Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0F, 96.0F)
            Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
            Me.Font = New System.Drawing.Font("Segoe UI", 9.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

            Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
            Me.Icon = Icon.FromHandle((New System.Drawing.Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))).GetHicon())
            Me.Text = Title

            ' Main TableLayoutPanel (4 rows)
            Dim tlpMain As New System.Windows.Forms.TableLayoutPanel() With {
                                            .Dock = System.Windows.Forms.DockStyle.Fill,
                                            .ColumnCount = 1,
                                            .RowCount = 4
                                        }
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Row 1: Label
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0F)) ' Row 2: ListBox
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Row 3: Checkbox
            tlpMain.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.AutoSize))    ' Row 4: Buttons

            ' Row 1: Label (shrinks & grows, 20px Padding)
            lblTitle = New System.Windows.Forms.Label() With {
                                            .Text = ListType,
                                            .AutoSize = True,
                                            .Dock = System.Windows.Forms.DockStyle.Fill,
                                            .Margin = New System.Windows.Forms.Padding(20, 20, 20, 0)
                                        }
            tlpMain.Controls.Add(lblTitle, 0, 0)

            ' Row 2: ListBox (shrinks & grows, 20px Padding)
            lstModels = New System.Windows.Forms.ListBox() With {
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .Margin = New System.Windows.Forms.Padding(20)
                                    }
            tlpMain.Controls.Add(lstModels, 0, 1)


            ' Row 3: Checkbox (grows but not shrink, 20px Padding)
            chkReset = New System.Windows.Forms.CheckBox() With {
                                        .Text = OptionText,
                                        .Checked = OptionChecked,
                                        .AutoSize = True,
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .Margin = New System.Windows.Forms.Padding(20, 0, 20, 0)
                                    }

            If OptionText <> "" Then
                tlpMain.Controls.Add(chkReset, 0, 2)
            End If

            ' Row 4: Buttons (left-to-right, grows but not shrink, 20px Padding)
            Dim flpButtons As New System.Windows.Forms.FlowLayoutPanel() With {
                                        .Dock = System.Windows.Forms.DockStyle.Fill,
                                        .FlowDirection = System.Windows.Forms.FlowDirection.LeftToRight,
                                        .AutoSize = True,
                                        .Margin = New System.Windows.Forms.Padding(20)
                                    }
            btnOK = New System.Windows.Forms.Button() With {
                                        .Text = "OK",
                                        .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                                    }
            AddHandler btnOK.Click, AddressOf btnOK_Click
            flpButtons.Controls.Add(btnOK)

            btnCancel = New System.Windows.Forms.Button() With {
                                        .Text = "Cancel",
                                        .Padding = New System.Windows.Forms.Padding(10, 5, 10, 5),
                                        .AutoSize = True,
                                        .AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
                                    }
            AddHandler btnCancel.Click, AddressOf btnCancel_Click
            flpButtons.Controls.Add(btnCancel)

            tlpMain.Controls.Add(flpButtons, 0, 3)

            Me.Controls.Add(tlpMain)
            Me.AcceptButton = btnOK
            Me.CancelButton = btnCancel

            ' Load models from INI file
            alternativeModels = LoadAlternativeModels(iniFilePath, context, Title)
            If UseCase = 1 Then
                lstModels.Items.Add("Default = " & context.INI_Model_2)
                hasDefaultEntry = True
            Else
                hasDefaultEntry = False
            End If
            For Each model In alternativeModels
                Dim displayText As String = If(String.IsNullOrEmpty(model.ModelDescription), model.Model, model.ModelDescription)
                lstModels.Items.Add(displayText)
            Next
            lstModels.SelectedIndex = 0
            AddHandler lstModels.DoubleClick, AddressOf lstModels_DoubleClick

            Me.ClientSize = New System.Drawing.Size(580, 450)
            Me.MinimumSize = Me.Size
        End Sub

        ''' <summary>
        ''' Handles double-click on the model list by applying the current selection.
        ''' </summary>
        Private Sub lstModels_DoubleClick(sender As Object, e As System.EventArgs)
            If lstModels.SelectedIndex >= 0 Then
                btnOK.PerformClick()
            End If
        End Sub


        ''' <summary>
        ''' Applies additional DPI scaling after the window handle is created.
        ''' </summary>
        Protected Overrides Sub OnHandleCreated(e As System.EventArgs)
            MyBase.OnHandleCreated(e)
            Dim dpiScale As Single = Me.DeviceDpi / 96.0F
            If dpiScale <> 1.0F Then
                Me.Scale(New System.Drawing.SizeF(dpiScale, dpiScale))
            End If
        End Sub



        ''' <summary>
        ''' Applies the current selection, updates shared flags, and closes the dialog with <see cref="DialogResult.OK"/>.
        ''' </summary>
        Private Sub btnOK_Click(sender As Object, e As EventArgs)
            Try
                If hasDefaultEntry AndAlso lstModels.SelectedIndex = 0 Then
                    UseDefault = True
                Else
                    UseDefault = False
                    ' adjust the index offset by 1 if there was a default entry
                    Dim offset As Integer = If(hasDefaultEntry, 1, 0)
                    Dim idx As Integer = lstModels.SelectedIndex - offset
                    If idx >= 0 AndAlso idx < alternativeModels.Count Then
                        SelectedModel = alternativeModels(idx)
                    End If
                End If

                ' If the checkbox is unchecked and a non-default model is selected, set OriginalConfigurationLoaded to False.
                If chkReset IsNot Nothing Then
                    If Not chkReset.Checked AndAlso Not UseDefault Then
                        originalConfigLoaded = False
                    End If
                    OptionChecked = chkReset.Checked
                Else
                    OptionChecked = True
                    If Not UseDefault Then
                        originalConfigLoaded = False
                    End If
                End If

                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As System.Exception
                MessageBox.Show("Error processing selection: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Cancels the dialog and closes it with <see cref="DialogResult.Cancel"/>.
        ''' </summary>
        Private Sub btnCancel_Click(sender As Object, e As EventArgs)
            Me.DialogResult = DialogResult.Cancel
            Me.Close()
        End Sub

    End Class

End Namespace