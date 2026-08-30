' Part of "Red Ink for Outlook"
' Copyright (c) LawDigital Ltd., Switzerland. All rights reserved.

' =============================================================================
' File: ReviewChangesDialog.vb
' Purpose: Interactive "track changes" style review dialog for Outlook.
'          Builds a word-level diff between the original selection and the LLM
'          corrected text using DiffPlex, renders it as a single RTF blob in a
'          RichTextBox, and lets the user click each highlighted token to
'          accept/reject it. The lower pane shows a live Markdown-rendered
'          preview of the resulting text.
' =============================================================================

Option Explicit On
Option Strict Off

Imports System.Collections.Generic
Imports System.Drawing
Imports System.Globalization
Imports System.Text
Imports System.Windows.Forms
Imports DiffPlex
Imports DiffPlex.DiffBuilder
Imports DiffPlex.DiffBuilder.Model
Imports SharedLibrary.SharedLibrary
Imports SharedLibrary.SharedLibrary.SharedMethods

Public Class ReviewChangesDialog
    Inherits System.Windows.Forms.Form

    Private Enum SegmentKind
        Equal
        Insertion
        Deletion
    End Enum

    Private Class Segment
        Public Kind As SegmentKind
        Public Text As String          ' visible text, incl. trailing space
        Public Accepted As Boolean = True
        Public VisibleStart As Integer ' index in plain text shown in RichTextBox
        Public VisibleLength As Integer
    End Class

    Private ReadOnly _segments As New List(Of Segment)()

    Private _diffBox As System.Windows.Forms.RichTextBox
    Private _previewBox As System.Windows.Forms.RichTextBox
    Private _statusLabel As System.Windows.Forms.Label
    Private _previewTimer As System.Windows.Forms.Timer
    Private _sentenceToggle As System.Windows.Forms.CheckBox

    ' Theme colors (kept in sync with RTF color table indices below)
    Private Shared ReadOnly ClrText As Color = Color.Black
    Private Shared ReadOnly ClrInsert As Color = Color.FromArgb(0, 80, 200)
    Private Shared ReadOnly ClrInsertReject As Color = Color.FromArgb(150, 150, 150)
    Private Shared ReadOnly ClrDelete As Color = Color.FromArgb(200, 30, 30)
    Private Shared ReadOnly BgInsert As Color = Color.FromArgb(225, 235, 255)
    Private Shared ReadOnly BgInsertReject As Color = Color.FromArgb(240, 240, 240)
    Private Shared ReadOnly BgDelete As Color = Color.FromArgb(255, 228, 228)
    Private Shared ReadOnly BgDeleteReject As Color = Color.FromArgb(255, 248, 205)

    Public Property ReviewedText As String = String.Empty

    Public Sub New(originalText As String, suggestedText As String)
        BuildSegments(If(originalText, ""), If(suggestedText, ""))
        InitializeUI()
        RenderDiffFast()
        SchedulePreview(immediate:=True)
        UpdateStatus()

        ' Surface any same-process native host prompt (e.g. Word's "Save changes?")
        ' that would otherwise stay hidden behind this TopMost modal dialog.
        SharedMethods.AttachForeignForegroundWatchdog(Me)
        AddHandler Me.Shown, Sub(sender, e) SharedMethods.ForceDialogToForeground(Me)
    End Sub

    ' --------------------------------------------------------------- UI build

    Private Sub InitializeUI()

        Me.Text = "Review changes"
        Me.ShowInTaskbar = False
        Me.KeyPreview = True
        Me.Font = New Font("Segoe UI", 9.0F, FontStyle.Regular, GraphicsUnit.Point)
        Me.AutoScaleMode = AutoScaleMode.Dpi
        Me.AutoScaleDimensions = New SizeF(96.0F, 96.0F)
        Me.FormBorderStyle = FormBorderStyle.Sizable
        Me.MaximizeBox = True
        Me.MinimizeBox = True
        Me.TopMost = True
        Me.MinimumSize = New Size(820, 540)

        ' Size and center relative to the screen: 60% height, 16:9 aspect ratio
        Dim wa As Rectangle = Screen.PrimaryScreen.WorkingArea
        Dim h As Integer = Math.Max(Me.MinimumSize.Height, CInt(wa.Height * 0.6))
        Dim w As Integer = Math.Max(Me.MinimumSize.Width, CInt(h * 16 / 9.0R))

        ' If the computed width exceeds the working area, cap width and recompute height from 16:9
        If w > wa.Width Then
            w = wa.Width
            h = Math.Max(Me.MinimumSize.Height, CInt(w * 9 / 16.0R))
        End If

        Me.StartPosition = FormStartPosition.Manual
        Me.Size = New Size(w, h)
        Me.Location = New Point(wa.Left + (wa.Width - w) \ 2, wa.Top + (wa.Height - h) \ 2)

        Try
            Dim bmp As New Bitmap(SharedMethods.GetLogoBitmap(SharedMethods.LogoType.Standard))
            Me.Icon = Icon.FromHandle(bmp.GetHicon())
        Catch
        End Try

        Dim root As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 3,
            .Padding = New Padding(16, 14, 16, 14)
        }
        root.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.Percent, 100))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        root.RowStyles.Add(New RowStyle(SizeType.AutoSize))

        ' --- Split: diff (top) / preview (bottom) ---
        Dim split As New SplitContainer() With {
            .Dock = DockStyle.Fill,
            .Orientation = Orientation.Horizontal,
            .SplitterWidth = 6,
            .Panel1MinSize = 160,
            .Panel2MinSize = 140
        }

        split.Panel1.Controls.Add(BuildDiffPanel())
        split.Panel2.Controls.Add(BuildPreviewPanel())

        AddHandler split.HandleCreated,
            Sub()
                Try
                    split.SplitterDistance = CInt(split.Height * 0.55)
                Catch
                End Try
            End Sub

        root.Controls.Add(split, 0, 0)

        ' --- Legend + status (single row, AutoSize) ---
        root.Controls.Add(BuildLegendAndStatus(), 0, 1)

        ' --- Buttons row ---
        root.Controls.Add(BuildButtonRow(), 0, 2)

        Me.Controls.Add(root)

        AddHandler Me.KeyDown,
            Sub(sender, e)
                If e.KeyCode = Keys.Escape Then
                    Me.DialogResult = DialogResult.Cancel
                    Me.Close()
                    e.SuppressKeyPress = True
                ElseIf e.Control AndAlso e.KeyCode = Keys.A Then
                    SetAll(True)
                    e.SuppressKeyPress = True
                ElseIf e.Control AndAlso e.KeyCode = Keys.R Then
                    SetAll(False)
                    e.SuppressKeyPress = True
                End If
            End Sub

        AddHandler Me.Shown,
            Sub()
                Me.TopMost = False
                Me.TopMost = True
                Me.Activate()
                Me.BringToFront()
            End Sub

        ' Debounced preview rebuild
        _previewTimer = New System.Windows.Forms.Timer() With {.Interval = 90}
        AddHandler _previewTimer.Tick,
            Sub()
                _previewTimer.Stop()
                RenderPreviewMarkdown()
            End Sub
    End Sub

    Private Function BuildDiffPanel() As Control
        Dim panel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2
        }
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        Dim caption As New Label() With {
            .Text = "Suggested changes — click any highlighted word to toggle accept / reject",
            .AutoSize = True,
            .Margin = New Padding(2, 0, 2, 6),
            .Font = New Font("Segoe UI Semibold", 9.0F)
        }

        _diffBox = New System.Windows.Forms.RichTextBox() With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.White,
            .HideSelection = True,
            .DetectUrls = False,
            .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point),
            .ScrollBars = RichTextBoxScrollBars.Vertical,
            .WordWrap = True
        }
        AddHandler _diffBox.MouseClick, AddressOf DiffBox_MouseClick
        AddHandler _diffBox.MouseMove, AddressOf DiffBox_MouseMove

        panel.Controls.Add(caption, 0, 0)
        panel.Controls.Add(_diffBox, 0, 1)
        Return panel
    End Function

    Private Function BuildPreviewPanel() As Control
        Dim panel As New TableLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .ColumnCount = 1,
            .RowCount = 2
        }
        panel.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))
        panel.RowStyles.Add(New RowStyle(SizeType.AutoSize))
        panel.RowStyles.Add(New RowStyle(SizeType.Percent, 100))

        Dim caption As New Label() With {
            .Text = "Final preview — this is the text that will replace your selection",
            .AutoSize = True,
            .Margin = New Padding(2, 6, 2, 6),
            .Font = New Font("Segoe UI Semibold", 9.0F)
        }

        _previewBox = New System.Windows.Forms.RichTextBox() With {
            .Dock = DockStyle.Fill,
            .ReadOnly = True,
            .BorderStyle = BorderStyle.FixedSingle,
            .BackColor = Color.FromArgb(250, 250, 250),
            .HideSelection = True,
            .DetectUrls = False,
            .Font = New Font("Segoe UI", 10.5F, FontStyle.Regular, GraphicsUnit.Point),
            .ScrollBars = RichTextBoxScrollBars.Vertical,
            .WordWrap = True
        }

        panel.Controls.Add(caption, 0, 0)
        panel.Controls.Add(_previewBox, 0, 1)
        Return panel
    End Function

    Private Function BuildLegendAndStatus() As Control
        Dim row As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.LeftToRight,
            .AutoSize = True,
            .WrapContents = False,
            .Margin = New Padding(0, 8, 0, 6)
        }

        row.Controls.Add(MakeLegendItem(BgInsert, ClrInsert, "Added", underline:=True))
        row.Controls.Add(MakeLegendItem(BgDelete, ClrDelete, "Removed", strikeout:=True))
        row.Controls.Add(MakeLegendItem(BgDeleteReject, ClrText, "Kept (deletion rejected)"))
        row.Controls.Add(MakeLegendItem(BgInsertReject, ClrInsertReject, "Skipped (addition rejected)", strikeout:=True))

        _sentenceToggle = New CheckBox() With {
            .AutoSize = True,
            .Text = "Toggle whole sentence",
            .Checked = False,
            .Font = New Font("Segoe UI", 9.0F),
            .Margin = New Padding(24, 4, 0, 0),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        row.Controls.Add(_sentenceToggle)

        _statusLabel = New Label() With {
            .AutoSize = True,
            .ForeColor = Color.FromArgb(50, 50, 50),
            .Font = New Font("Segoe UI Semibold", 9.0F),
            .Margin = New Padding(24, 6, 0, 0),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        row.Controls.Add(_statusLabel)

        Return row
    End Function

    Private Function MakeLegendItem(background As Color, foreground As Color, text As String,
                                    Optional underline As Boolean = False,
                                    Optional strikeout As Boolean = False) As Control
        Dim style As FontStyle = FontStyle.Regular
        If underline Then style = style Or FontStyle.Underline
        If strikeout Then style = style Or FontStyle.Strikeout

        Dim swatch As New Label() With {
            .AutoSize = False,
            .Width = 18,
            .Height = 18,
            .BackColor = background,
            .Margin = New Padding(0, 4, 6, 0),
            .BorderStyle = BorderStyle.FixedSingle
        }

        Dim label As New Label() With {
            .AutoSize = True,
            .Text = text,
            .ForeColor = foreground,
            .Font = New Font("Segoe UI", 9.0F, style),
            .Margin = New Padding(0, 6, 18, 0),
            .TextAlign = ContentAlignment.MiddleLeft
        }

        Dim wrap As New FlowLayoutPanel() With {
            .AutoSize = True,
            .FlowDirection = FlowDirection.LeftToRight,
            .WrapContents = False,
            .Margin = New Padding(0)
        }
        wrap.Controls.Add(swatch)
        wrap.Controls.Add(label)
        Return wrap
    End Function

    Private Function BuildButtonRow() As Control
        Dim row As New FlowLayoutPanel() With {
            .Dock = DockStyle.Fill,
            .FlowDirection = FlowDirection.RightToLeft,
            .AutoSize = True,
            .WrapContents = False,
            .Margin = New Padding(0, 6, 0, 0)
        }

        Dim cancelBtn As Button = MakeButton("Cancel", primary:=False)
        Dim insertBtn As Button = MakeButton("Insert reviewed text", primary:=True)
        Dim rejectAllBtn As Button = MakeButton("Reject all", primary:=False)
        Dim acceptAllBtn As Button = MakeButton("Accept all", primary:=False)

        AddHandler cancelBtn.Click,
            Sub()
                Me.DialogResult = DialogResult.Cancel
                Me.Close()
            End Sub

        AddHandler insertBtn.Click,
            Sub()
                Me.ReviewedText = BuildReviewedText()
                Me.DialogResult = DialogResult.OK
                Me.Close()
            End Sub

        AddHandler acceptAllBtn.Click, Sub() SetAll(True)
        AddHandler rejectAllBtn.Click, Sub() SetAll(False)

        row.Controls.Add(cancelBtn)
        row.Controls.Add(insertBtn)
        row.Controls.Add(rejectAllBtn)
        row.Controls.Add(acceptAllBtn)

        Me.AcceptButton = insertBtn
        Me.CancelButton = cancelBtn

        Return row
    End Function

    Private Function MakeButton(text As String, primary As Boolean) As Button
        Return New Button() With {
            .Text = text,
            .AutoSize = True,
            .AutoSizeMode = AutoSizeMode.GrowAndShrink,
            .Padding = New Padding(16, 6, 16, 6),
            .Margin = New Padding(8, 0, 0, 0),
            .UseVisualStyleBackColor = True,
            .Font = New Font("Segoe UI", 9.0F, If(primary, FontStyle.Bold, FontStyle.Regular))
        }
    End Function

    ' ----------------------------------------------------------- Diff build

    Private Sub BuildSegments(originalText As String, suggestedText As String)

        Dim builder As New InlineDiffBuilder(New Differ())

        Dim aTokens As List(Of String) = TokenizePreservingWhitespace(originalText)
        Dim bTokens As List(Of String) = TokenizePreservingWhitespace(suggestedText)

        Dim aWords As String = String.Join(Environment.NewLine, aTokens)
        Dim bWords As String = String.Join(Environment.NewLine, bTokens)

        Dim model As DiffPaneModel = builder.BuildDiffModel(aWords, bWords)

        For Each line As DiffPiece In model.Lines

            Dim raw As String = If(line.Text, "")
            If raw.Length = 0 Then Continue For

            Dim restored As String = raw _
            .Replace("{br}", vbCrLf) _
            .Replace("{vbCrLf}", vbCrLf) _
            .Replace("{vbCr}", vbCrLf) _
            .Replace("{vbLf}", vbCrLf)

            If restored.Length = 0 Then Continue For

            Dim displayText As String = restored

            Dim kind As SegmentKind
            Select Case line.Type
                Case ChangeType.Inserted
                    kind = SegmentKind.Insertion
                Case ChangeType.Deleted
                    kind = SegmentKind.Deletion
                Case Else
                    kind = SegmentKind.Equal
            End Select

            _segments.Add(New Segment() With {
            .Kind = kind,
            .Text = displayText,
            .Accepted = True
        })
        Next

        ConsolidateEquivalentLineBreakChanges()
    End Sub

    Private Shared Function TokenizePreservingWhitespace(s As System.String) As System.Collections.Generic.List(Of System.String)
        Dim tokens As New System.Collections.Generic.List(Of System.String)()
        If s Is Nothing Then Return tokens

        Dim normalized As System.String =
            s.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf)

        Dim lines As System.String() = normalized.Split(ControlChars.Lf)

        For lineIndex As System.Int32 = 0 To lines.Length - 1
            TokenizeLinePreservingWhitespace(lines(lineIndex), tokens)
            If lineIndex < lines.Length - 1 Then tokens.Add("{br}")
        Next

        Return tokens
    End Function

    Private Shared Sub TokenizeLinePreservingWhitespace(
        line As System.String,
        tokens As System.Collections.Generic.List(Of System.String))

        If tokens Is Nothing Then Return
        If line Is Nothing Then line = System.String.Empty

        Dim offset As System.Int32 = 0

        ' Treat the complete Markdown list prefix as one structural token. This prevents
        ' ReviewChangesDialog from allowing indentation and marker changes to be accepted
        ' independently, which can turn a valid nested list into a code block or flatten it.
        Dim listPrefixMatch As System.Text.RegularExpressions.Match =
            System.Text.RegularExpressions.Regex.Match(
                line,
                "^([ \t]*)(?:[-+*]|\d+[\.)])([ \t]+)",
                System.Text.RegularExpressions.RegexOptions.CultureInvariant)

        If listPrefixMatch.Success AndAlso listPrefixMatch.Length > 0 Then
            tokens.Add(listPrefixMatch.Value)
            offset = listPrefixMatch.Length
        End If

        Dim i As System.Int32 = offset

        While i < line.Length
            Dim ch As System.Char = line(i)

            If ch = ControlChars.Tab Then
                tokens.Add(ControlChars.Tab)
                i += 1

            ElseIf ch = " "c Then
                Dim start As System.Int32 = i
                While i < line.Length AndAlso line(i) = " "c
                    i += 1
                End While
                tokens.Add(line.Substring(start, i - start))

            Else
                Dim start As System.Int32 = i
                While i < line.Length AndAlso
                      line(i) <> " "c AndAlso
                      line(i) <> ControlChars.Tab
                    i += 1
                End While
                tokens.Add(line.Substring(start, i - start))
            End If
        End While
    End Sub


    Private Sub ConsolidateEquivalentLineBreakChanges()
        If _segments.Count < 2 Then Return

        Dim consolidated As New List(Of Segment)(_segments.Count)
        Dim i As Integer = 0

        While i < _segments.Count
            If i < _segments.Count - 1 AndAlso IsEquivalentLineBreakChange(_segments(i), _segments(i + 1)) Then
                consolidated.Add(New Segment() With {
                .Kind = SegmentKind.Equal,
                .Text = CanonicalLineBreakText(_segments(i).Text),
                .Accepted = True
            })
                i += 2
            Else
                consolidated.Add(_segments(i))
                i += 1
            End If
        End While

        _segments.Clear()
        _segments.AddRange(consolidated)
    End Sub

    Private Shared Function IsEquivalentLineBreakChange(first As Segment, second As Segment) As Boolean
        If first Is Nothing OrElse second Is Nothing Then Return False

        Dim isOppositeChange As Boolean =
        (first.Kind = SegmentKind.Deletion AndAlso second.Kind = SegmentKind.Insertion) OrElse
        (first.Kind = SegmentKind.Insertion AndAlso second.Kind = SegmentKind.Deletion)

        If Not isOppositeChange Then Return False
        If Not IsLineBreakOnly(first.Text) OrElse Not IsLineBreakOnly(second.Text) Then Return False

        Return CountLineBreaks(first.Text) = CountLineBreaks(second.Text)
    End Function

    Private Shared Function IsLineBreakOnly(text As String) As Boolean
        If String.IsNullOrEmpty(text) Then Return False
        If CountLineBreaks(text) = 0 Then Return False

        For Each ch As Char In text
            Select Case ch
                Case " "c, ControlChars.Cr, ControlChars.Lf
                    Continue For
                Case Else
                    Return False
            End Select
        Next

        Return True
    End Function

    Private Shared Function CountLineBreaks(text As String) As Integer
        If text Is Nothing Then Return 0

        Dim count As Integer = 0
        Dim i As Integer = 0

        While i < text.Length
            If text(i) = ControlChars.Cr Then
                count += 1
                If i < text.Length - 1 AndAlso text(i + 1) = ControlChars.Lf Then
                    i += 2
                Else
                    i += 1
                End If
            ElseIf text(i) = ControlChars.Lf Then
                count += 1
                i += 1
            Else
                i += 1
            End If
        End While

        Return count
    End Function

    Private Shared Function CanonicalLineBreakText(text As String) As String
        Dim count As Integer = CountLineBreaks(text)
        If count <= 0 Then Return ""

        Dim sb As New StringBuilder(count * 2)
        For i As Integer = 1 To count
            sb.Append(vbCrLf)
        Next

        Return sb.ToString()
    End Function

    ' ----------------------------------------------------------- Fast RTF

    ' Color table indices used in RTF below (1-based as per RTF spec; index 0 = auto/default).
    Private Const RTF_FG_TEXT As Integer = 1          ' black
    Private Const RTF_FG_INSERT As Integer = 2        ' blue
    Private Const RTF_FG_INSERT_REJ As Integer = 3    ' gray
    Private Const RTF_FG_DELETE As Integer = 4        ' red
    Private Const RTF_BG_WHITE As Integer = 5
    Private Const RTF_BG_INSERT As Integer = 6
    Private Const RTF_BG_INSERT_REJ As Integer = 7
    Private Const RTF_BG_DELETE As Integer = 8
    Private Const RTF_BG_DELETE_REJ As Integer = 9

    Private Shared Function ColorEntry(c As Color) As String
        Return "\red" & c.R.ToString(CultureInfo.InvariantCulture) &
               "\green" & c.G.ToString(CultureInfo.InvariantCulture) &
               "\blue" & c.B.ToString(CultureInfo.InvariantCulture) & ";"
    End Function

    Private Sub RenderDiffFast()

        Dim sb As New StringBuilder(_segments.Count * 24 + 512)

        ' RTF header + font table + color table
        sb.Append("{\rtf1\ansi\ansicpg1252\deff0")
        sb.Append("{\fonttbl{\f0\fnil\fcharset0 Segoe UI;}}")
        sb.Append("{\colortbl ;")
        sb.Append(ColorEntry(ClrText))
        sb.Append(ColorEntry(ClrInsert))
        sb.Append(ColorEntry(ClrInsertReject))
        sb.Append(ColorEntry(ClrDelete))
        sb.Append(ColorEntry(Color.White))
        sb.Append(ColorEntry(BgInsert))
        sb.Append(ColorEntry(BgInsertReject))
        sb.Append(ColorEntry(BgDelete))
        sb.Append(ColorEntry(BgDeleteReject))
        sb.Append("}")
        sb.Append("\fs21 ")  ' 10.5 pt ≈ \fs21

        Dim visibleIndex As Integer = 0

        For Each seg In _segments

            Dim fg As Integer = RTF_FG_TEXT
            Dim bg As Integer = RTF_BG_WHITE
            Dim ul As Boolean = False
            Dim strike As Boolean = False

            Select Case seg.Kind
                Case SegmentKind.Equal
                    fg = RTF_FG_TEXT : bg = RTF_BG_WHITE

                Case SegmentKind.Insertion
                    If seg.Accepted Then
                        fg = RTF_FG_INSERT : bg = RTF_BG_INSERT : ul = True
                    Else
                        fg = RTF_FG_INSERT_REJ : bg = RTF_BG_INSERT_REJ : strike = True
                    End If

                Case SegmentKind.Deletion
                    If seg.Accepted Then
                        fg = RTF_FG_DELETE : bg = RTF_BG_DELETE : strike = True
                    Else
                        fg = RTF_FG_TEXT : bg = RTF_BG_DELETE_REJ
                    End If
            End Select

            seg.VisibleStart = visibleIndex

            ' Open formatting group
            sb.Append("{\cf").Append(fg.ToString(CultureInfo.InvariantCulture))
            sb.Append("\highlight").Append(bg.ToString(CultureInfo.InvariantCulture))
            If ul Then sb.Append("\ul")
            If strike Then sb.Append("\strike")
            sb.Append(" ")

            AppendEscapedRtf(sb, seg.Text)

            sb.Append("}")

            ' Visible length equals raw chars, with CRLF counted as 1 (the RichTextBox normalizes)
            seg.VisibleLength = CountVisibleChars(seg.Text)
            visibleIndex += seg.VisibleLength
        Next

        sb.Append("}")

        _diffBox.SuspendLayout()
        Try
            _diffBox.Rtf = sb.ToString()
            _diffBox.Select(0, 0)
        Finally
            _diffBox.ResumeLayout()
        End Try
    End Sub

    Private Shared Sub AppendEscapedRtf(sb As StringBuilder, text As String)
        For Each ch As Char In text
            Select Case ch
                Case "\"c
                    sb.Append("\\")
                Case "{"c
                    sb.Append("\{")
                Case "}"c
                    sb.Append("\}")
                Case ControlChars.Cr
                    ' Drop bare CR; CRLF emits \par via the LF case below
                Case ControlChars.Lf
                    sb.Append("\par ")
                Case ControlChars.Tab
                    sb.Append("\tab ")
                Case Else
                    Dim code As Integer = AscW(ch)
                    If code < 128 Then
                        sb.Append(ch)
                    Else
                        sb.Append("\u").Append(code.ToString(CultureInfo.InvariantCulture)).Append("?")
                    End If
            End Select
        Next
    End Sub

    Private Shared Function CountVisibleChars(text As String) As Integer
        ' RichTextBox treats CRLF as a single newline in its text indexing.
        Dim n As Integer = 0
        Dim i As Integer = 0
        While i < text.Length
            If i < text.Length - 1 AndAlso text(i) = ControlChars.Cr AndAlso text(i + 1) = ControlChars.Lf Then
                n += 1
                i += 2
            Else
                n += 1
                i += 1
            End If
        End While
        Return n
    End Function

    ' ----------------------------------------------------------- Preview

    Private Sub SchedulePreview(Optional immediate As Boolean = False)
        If _previewTimer Is Nothing Then Return
        _previewTimer.Stop()
        If immediate Then
            RenderPreviewMarkdown()
        Else
            _previewTimer.Start()
        End If
    End Sub

    Private Sub RenderPreviewMarkdown()
        Dim md As String = BuildReviewedText()
        Try
            If String.IsNullOrWhiteSpace(md) Then
                _previewBox.Clear()
                Return
            End If

            Dim rtf As String = MarkdownToRtfConverter.Convert(md, preserveSquareBracketLiterals:=True)
            _previewBox.Rtf = rtf
            _previewBox.Select(0, 0)
        Catch
            ' Fallback to plain text if conversion fails
            _previewBox.Clear()
            _previewBox.Text = md
        End Try
    End Sub

    ' ----------------------------------------------------------- Interaction

    Private Sub DiffBox_MouseClick(sender As Object, e As MouseEventArgs)
        If e.Button <> MouseButtons.Left Then Return

        Dim idx As Integer = _diffBox.GetCharIndexFromPosition(e.Location)
        If idx < 0 Then Return

        Dim seg As Segment = FindSegmentAt(idx)
        If seg Is Nothing OrElse seg.Kind = SegmentKind.Equal Then Return

        Dim newState As Boolean = Not seg.Accepted

        If _sentenceToggle IsNot Nothing AndAlso _sentenceToggle.Checked Then
            ApplyStateToSentence(seg, newState)
        Else
            seg.Accepted = newState
        End If

        ' Re-render: this is a single Rtf assignment, very fast even for large diffs
        Dim caret As Integer = idx
        RenderDiffFast()
        SchedulePreview()
        UpdateStatus()

        ' Restore approximate caret/visible position
        If caret >= 0 AndAlso caret <= _diffBox.TextLength Then
            _diffBox.Select(caret, 0)
            _diffBox.ScrollToCaret()
        End If
    End Sub

    Private Sub DiffBox_MouseMove(sender As Object, e As MouseEventArgs)
        Dim idx As Integer = _diffBox.GetCharIndexFromPosition(e.Location)
        Dim seg As Segment = FindSegmentAt(idx)
        Dim wantHand As Boolean = (seg IsNot Nothing AndAlso seg.Kind <> SegmentKind.Equal)
        Dim target As Cursor = If(wantHand, Cursors.Hand, Cursors.IBeam)
        If _diffBox.Cursor IsNot target Then _diffBox.Cursor = target
    End Sub

    Private Function FindSegmentAt(charIndex As Integer) As Segment
        If charIndex < 0 Then Return Nothing
        For Each seg In _segments
            Dim len As Integer = Math.Max(seg.VisibleLength, 1)
            If charIndex >= seg.VisibleStart AndAlso charIndex < seg.VisibleStart + len Then
                Return seg
            End If
        Next
        Return Nothing
    End Function

    ' Toggles every change segment (insertion/deletion) that belongs to the same
    ' sentence as the clicked segment to a single, uniform accept/reject state, so
    ' a series of consecutive changes can be handled with one click.
    Private Sub ApplyStateToSentence(clicked As Segment, newState As Boolean)
        Dim center As Integer = _segments.IndexOf(clicked)
        If center < 0 Then
            clicked.Accepted = newState
            Return
        End If

        ' Walk left: the sentence starts right after the previous terminator.
        Dim first As Integer = center
        While first > 0 AndAlso Not IsSentenceTerminator(_segments(first - 1))
            first -= 1
        End While

        ' Walk right: the sentence ends at (and includes) the next terminator.
        Dim last As Integer = center
        While last < _segments.Count - 1 AndAlso Not IsSentenceTerminator(_segments(last))
            last += 1
        End While

        For i As Integer = first To last
            If _segments(i).Kind <> SegmentKind.Equal Then
                _segments(i).Accepted = newState
            End If
        Next
    End Sub

    ' A segment terminates a sentence when its visible text ends with sentence
    ' punctuation or contains a line break.
    Private Shared Function IsSentenceTerminator(seg As Segment) As Boolean
        If seg Is Nothing Then Return False
        If CountLineBreaks(seg.Text) > 0 Then Return True

        Dim trimmed As String = If(seg.Text, "").TrimEnd(" "c, ControlChars.Tab)
        If trimmed.Length = 0 Then Return False

        Select Case trimmed(trimmed.Length - 1)
            Case "."c, "!"c, "?"c, "…"c
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Sub SetAll(accept As Boolean)
        For Each seg In _segments
            If seg.Kind <> SegmentKind.Equal Then seg.Accepted = accept
        Next
        RenderDiffFast()
        SchedulePreview()
        UpdateStatus()
    End Sub

    ' ----------------------------------------------------------- Result text

    Private Shared Function IsIncludedInResult(seg As Segment) As Boolean
        Select Case seg.Kind
            Case SegmentKind.Equal : Return True
            Case SegmentKind.Insertion : Return seg.Accepted
            Case SegmentKind.Deletion : Return Not seg.Accepted
            Case Else : Return False
        End Select
    End Function

    Private Function BuildReviewedText() As String
        Dim sb As New StringBuilder()
        For Each seg In _segments
            If IsIncludedInResult(seg) Then sb.Append(seg.Text)
        Next
        Return sb.ToString()
    End Function

    Private Sub UpdateStatus()
        Dim total As Integer = 0
        Dim accepted As Integer = 0
        For Each seg In _segments
            If seg.Kind <> SegmentKind.Equal Then
                total += 1
                If seg.Accepted Then accepted += 1
            End If
        Next
        If total = 0 Then
            _statusLabel.Text = "No differences detected"
        Else
            _statusLabel.Text = $"{accepted} of {total} suggested changes accepted"
        End If
    End Sub

End Class
