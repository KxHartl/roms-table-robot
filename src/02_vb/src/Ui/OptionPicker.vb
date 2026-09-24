Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Windows.Forms

' One entry of a drop-down: what the user reads, what the model stores, and what the
' choice actually does.
Public Class ChoiceItem

    Public ReadOnly Display As String
    Public ReadOnly Value As String
    Public ReadOnly Hint As String

    Public Sub New(ByVal display_ As String, ByVal value_ As String, ByVal hint_ As String)
        Display = display_
        Value = value_
        Hint = hint_
    End Sub

    Public Overrides Function ToString() As String
        Return Display
    End Function

End Class

' A drop-down that only offers valid catalogue values (typing is impossible) and prints
' the consequence of the selected option underneath.
Public Class OptionPicker
    Inherits UserControl

    Private ReadOnly _caption As New Label()
    Private ReadOnly _combo As New ComboBox()
    Private ReadOnly _hint As New Label()

    Public Property SketchKey As String = ""

    Public Event SelectionChanged As EventHandler
    Public Event Activated As EventHandler

    Public Sub New(ByVal caption As String, ByVal sketchKey_ As String, ByVal items As ChoiceItem())
        SketchKey = sketchKey_
        Me.Height = 82
        Me.Margin = New Padding(0, 0, 0, 10)
        Me.BackColor = Color.Transparent

        _caption.AutoSize = True
        _caption.Font = Theme.BodyBold
        _caption.ForeColor = Theme.Ink
        _caption.Location = New Point(2, 2)
        _caption.Text = caption

        _combo.DropDownStyle = ComboBoxStyle.DropDownList
        _combo.Font = Theme.Body
        _combo.FlatStyle = FlatStyle.Standard
        _combo.Location = New Point(2, 22)
        _combo.Height = 24
        _combo.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        _combo.Items.AddRange(items)

        _hint.AutoSize = False
        _hint.Font = Theme.Small
        _hint.ForeColor = Theme.Muted
        _hint.Location = New Point(2, 50)
        _hint.Height = 30
        _hint.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Me.Controls.AddRange(New Control() {_caption, _combo, _hint})

        AddHandler _combo.SelectedIndexChanged, AddressOf OnPicked
        AddHandler _combo.Enter, AddressOf OnActivated
        AddHandler _combo.DropDown, AddressOf OnActivated
        ' Hovering is enough to switch the sketch - point at an input, see what it means.
        AddHandler Me.MouseEnter, AddressOf OnActivated
        AddHandler _caption.MouseEnter, AddressOf OnActivated
        AddHandler _hint.MouseEnter, AddressOf OnActivated
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        _combo.Width = Me.ClientSize.Width - 34
        _hint.Width = Me.ClientSize.Width - 6
    End Sub

    Private Sub OnActivated(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent Activated(Me, EventArgs.Empty)
    End Sub

    Private Sub OnPicked(ByVal sender As Object, ByVal e As EventArgs)
        Dim it As ChoiceItem = Selected()
        _hint.Text = If(it Is Nothing, "", it.Hint)
        Theme.Tips.SetToolTip(_combo, _hint.Text)
        RaiseEvent SelectionChanged(Me, EventArgs.Empty)
    End Sub

    Public Function Selected() As ChoiceItem
        If _combo.SelectedIndex < 0 Then Return Nothing
        Return DirectCast(_combo.SelectedItem, ChoiceItem)
    End Function

    ' Model value of the selection, or "" when nothing is selected.
    Public Property SelectedValue As String
        Get
            Dim it As ChoiceItem = Selected()
            If it Is Nothing Then Return ""
            Return it.Value
        End Get
        Set(ByVal v As String)
            For i As Integer = 0 To _combo.Items.Count - 1
                Dim it As ChoiceItem = DirectCast(_combo.Items(i), ChoiceItem)
                If String.Equals(it.Value, v, StringComparison.OrdinalIgnoreCase) Then
                    _combo.SelectedIndex = i
                    Return
                End If
            Next
            _combo.SelectedIndex = -1
        End Set
    End Property

    Public Sub SelectFirstIfEmpty()
        If _combo.SelectedIndex < 0 AndAlso _combo.Items.Count > 0 Then _combo.SelectedIndex = 0
    End Sub

End Class
