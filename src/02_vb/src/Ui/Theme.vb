Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Windows.Forms

' One place for colours, fonts and the few widgets that repeat, so the whole window
' looks like a single application instead of a pile of default controls.
Public Module Theme

    Public ReadOnly Background As Color = Color.FromArgb(243, 245, 248)
    Public ReadOnly Card As Color = Color.White
    Public ReadOnly Border As Color = Color.FromArgb(222, 226, 232)
    Public ReadOnly Ink As Color = Color.FromArgb(28, 32, 38)
    Public ReadOnly Muted As Color = Color.FromArgb(110, 118, 130)
    Public ReadOnly Accent As Color = Color.FromArgb(0, 96, 170)
    Public ReadOnly AccentSoft As Color = Color.FromArgb(226, 238, 249)
    Public ReadOnly Ok As Color = Color.FromArgb(21, 115, 71)
    Public ReadOnly Warn As Color = Color.FromArgb(176, 96, 12)
    Public ReadOnly Bad As Color = Color.FromArgb(180, 35, 24)
    Public ReadOnly Steel As Color = Color.FromArgb(150, 160, 172)
    Public ReadOnly SteelDark As Color = Color.FromArgb(96, 106, 120)
    Public ReadOnly PlateFill As Color = Color.FromArgb(206, 214, 224)

    Public ReadOnly Body As New Font("Segoe UI", 9.0!)
    Public ReadOnly BodyBold As New Font("Segoe UI", 9.0!, FontStyle.Bold)
    Public ReadOnly Small As New Font("Segoe UI", 8.0!)
    Public ReadOnly SmallItalic As New Font("Segoe UI", 8.0!, FontStyle.Italic)
    Public ReadOnly Heading As New Font("Segoe UI Semibold", 11.0!)
    Public ReadOnly Title As New Font("Segoe UI Semibold", 14.0!)

    Public ReadOnly Tips As New ToolTip() With {.AutoPopDelay = 20000, .InitialDelay = 350, .ReshowDelay = 120}

    Public Function SectionHeader(ByVal caption As String, ByVal subtitle As String, ByVal w As Integer) As Control
        Dim host As New Panel()
        host.Width = w
        host.Height = If(String.IsNullOrEmpty(subtitle), 30, 48)
        host.Margin = New Padding(0, 12, 0, 2)
        host.BackColor = Color.Transparent

        Dim t As New Label()
        t.AutoSize = True
        t.Font = Heading
        t.ForeColor = Ink
        t.Location = New Point(2, 4)
        t.Text = caption
        host.Controls.Add(t)

        If Not String.IsNullOrEmpty(subtitle) Then
            Dim s As New Label()
            s.AutoSize = False
            s.Font = Small
            s.ForeColor = Muted
            s.Location = New Point(2, 25)
            s.Size = New Size(w - 12, 20)
            s.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
            s.Text = subtitle
            host.Controls.Add(s)
        End If
        Return host
    End Function

    Public Function PrimaryButton(ByVal caption As String, ByVal w As Integer, ByVal h As Integer) As Button
        Dim b As New Button()
        b.Text = caption
        b.Size = New Size(w, h)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 0
        b.BackColor = Accent
        b.ForeColor = Color.White
        b.Font = BodyBold
        b.Cursor = Cursors.Hand
        Return b
    End Function

    Public Function QuietButton(ByVal caption As String, ByVal w As Integer, ByVal h As Integer) As Button
        Dim b As New Button()
        b.Text = caption
        b.Size = New Size(w, h)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderColor = Border
        b.FlatAppearance.BorderSize = 1
        b.BackColor = Card
        b.ForeColor = Ink
        b.Font = Body
        b.Cursor = Cursors.Hand
        Return b
    End Function

    Public Sub StretchChildren(ByVal flow As FlowLayoutPanel)
        Dim w As Integer = flow.ClientSize.Width - flow.Padding.Horizontal - 4
        If w < 220 Then w = 220
        For Each c As Control In flow.Controls
            c.Width = w - c.Margin.Horizontal
        Next
    End Sub

End Module

' Panel that paints its own hairline border - avoids the 3D look of BorderStyle.FixedSingle.
Public Class CardBox
    Inherits Panel

    Public Sub New()
        Me.BackColor = Theme.Card
    End Sub

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)
        Using pen As New Pen(Theme.Border)
            e.Graphics.DrawRectangle(pen, 0, 0, Me.Width - 1, Me.Height - 1)
        End Using
    End Sub

End Class
