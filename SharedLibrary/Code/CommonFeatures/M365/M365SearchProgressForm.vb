' Part of "Red Ink" (SharedLibrary)
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved. For license to use see https://redink.ai.
'
' =============================================================================
' File: M365SearchProgressForm.vb
' Purpose: Optional WinForms progress dialog for M365Service.SearchAsync.
'          Implements IProgress(Of M365SearchProgress) so callers can pass
'          the form directly:
'
'              Using f = New M365SearchProgressForm()
'                  f.Show(owner)
'                  Dim result = Await M365Service.SearchAsync(_context, "foo",
'                      M365SearchSources.All, progress:=f, ct:=f.CancellationToken)
'              End Using
'
'          Headless callers can ignore this form entirely and pass their own
'          IProgress(Of M365SearchProgress) (e.g. wired to ApDashboardLog).
'
' KEY ARCHITECTURE
'   • Surface: M365SearchProgressForm (host partial/UI).
'   • Entry points/helpers: Report, ApplyProgress, btnCancel_Click.
' =============================================================================

Option Explicit On
Option Strict On

Imports System.Threading
Imports System.Windows.Forms
Imports Microsoft.VisualBasic.Logging

Namespace SharedLibrary

    Public Class M365SearchProgressForm
        Inherits Form
        Implements IProgress(Of M365SearchProgress)

        Private ReadOnly _cts As New CancellationTokenSource()

        Public ReadOnly Property CancellationToken As CancellationToken
            Get
                Return _cts.Token
            End Get
        End Property

        Public Sub New()
            InitializeComponent()
        End Sub

        Public Sub Report(value As M365SearchProgress) Implements IProgress(Of M365SearchProgress).Report
            If InvokeRequired Then
                Try
                    BeginInvoke(New Action(Sub() ApplyProgress(value)))
                Catch
                End Try
            Else
                ApplyProgress(value)
            End If
        End Sub

        Private Sub ApplyProgress(p As M365SearchProgress)
            Dim line As String =
                $"[{DateTime.Now:HH:mm:ss}] {p.Stage}" &
                If(p.Source = M365SearchSources.None, "", $" — {p.Source}") &
                If(String.IsNullOrEmpty(p.Message), "", ": " & p.Message)
            lstLog.Items.Add(line)
            lstLog.TopIndex = lstLog.Items.Count - 1

            If p.Percent < 0 Then
                pb.Style = ProgressBarStyle.Marquee
            Else
                pb.Style = ProgressBarStyle.Continuous
                pb.Value = Math.Max(0, Math.Min(100, p.Percent))
            End If

            lblHits.Text = $"Hits so far: {p.HitsSoFar}"

            Select Case p.Stage
                Case M365ProgressStage.Completed
                    btnCancel.Text = "Close"
                    pb.Style = ProgressBarStyle.Continuous
                    pb.Value = 100
                Case M365ProgressStage.Failed, M365ProgressStage.Cancelled
                    btnCancel.Text = "Close"
            End Select
        End Sub

        Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
            If Not _cts.IsCancellationRequested Then _cts.Cancel()
            ' If search has already finished, the button label is "Close" — just dismiss.
            If btnCancel.Text = "Close" Then Close()
        End Sub

        Protected Overrides Sub OnFormClosed(e As FormClosedEventArgs)
            Try
                _cts.Cancel()
            Catch
            End Try
            MyBase.OnFormClosed(e)
        End Sub
    End Class

End Namespace