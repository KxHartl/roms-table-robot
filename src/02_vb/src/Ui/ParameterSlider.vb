Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Windows.Forms

' One numeric input: plain-language caption, one line of explanation, a slider bounded by
' the currently allowed range, and a spin box for an exact value.
'
' The slider cannot leave the range and the spin box clamps whatever is typed, so an
' impossible number never reaches CATIA. The range line under the slider always says what
' the limits are and why - the user is guided instead of corrected.
Public Class ParameterSlider
    Inherits UserControl

    Private ReadOnly _caption As New Label()
    Private ReadOnly _hint As New Label()
    Private ReadOnly _num As New NumericUpDown()
    Private ReadOnly _unit As New Label()
    Private ReadOnly _bar As New TrackBar()
    Private ReadOnly _range As New Label()

    Private _sync As Boolean = False
    Private _scale As Double = 1.0
    Private _available As Boolean = True
    Private _rangeText As String = ""
    Private _reason As String = ""

    Public Property SketchKey As String = ""

    Public Event ValueChanged As EventHandler
    Public Event Activated As EventHandler

    Public Sub New(ByVal caption As String, ByVal hint As String, ByVal unit As String,
                   ByVal decimals As Integer, ByVal sketchKey_ As String)
        _scale = Math.Pow(10, decimals)
        SketchKey = sketchKey_

        Me.Height = 96
        Me.Margin = New Padding(0, 0, 0, 10)
        Me.BackColor = Color.Transparent

        _caption.AutoSize = True
        _caption.Font = Theme.BodyBold
        _caption.ForeColor = Theme.Ink
        _caption.Location = New Point(2, 2)
        _caption.Text = caption

        _hint.AutoSize = False
        _hint.Font = Theme.Small
        _hint.ForeColor = Theme.Muted
        _hint.Location = New Point(2, 21)
        _hint.Height = 16
        _hint.Text = hint
        _hint.AutoEllipsis = True

        _num.Font = Theme.Body
        _num.DecimalPlaces = decimals
        _num.Increment = CDec(1 / _scale)
        _num.TextAlign = HorizontalAlignment.Right
        _num.Size = New Size(72, 24)
        _num.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        _unit.AutoSize = True
        _unit.Font = Theme.Small
        _unit.ForeColor = Theme.Muted
        _unit.Text = unit
        _unit.Anchor = AnchorStyles.Top Or AnchorStyles.Right

        _bar.AutoSize = False
        _bar.Height = 30
        _bar.Location = New Point(0, 40)
        _bar.TickStyle = TickStyle.None
        _bar.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        _range.AutoSize = False
        _range.Font = Theme.Small
        _range.ForeColor = Theme.Muted
        _range.Location = New Point(2, 70)
        _range.Height = 22
        _range.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

        Me.Controls.AddRange(New Control() {_caption, _hint, _num, _unit, _bar, _range})

        AddHandler _num.ValueChanged, AddressOf OnNumChanged
        AddHandler _bar.ValueChanged, AddressOf OnBarChanged
        AddHandler _num.Enter, AddressOf OnActivated
        AddHandler _bar.Enter, AddressOf OnActivated
        AddHandler _bar.MouseDown, AddressOf OnActivated
        ' Hovering is enough to switch the sketch - point at an input, see what it means.
        AddHandler Me.MouseEnter, AddressOf OnActivated
        AddHandler _caption.MouseEnter, AddressOf OnActivated
        AddHandler _hint.MouseEnter, AddressOf OnActivated
        AddHandler _bar.MouseEnter, AddressOf OnActivated
        AddHandler _range.MouseEnter, AddressOf OnActivated

        Theme.Tips.SetToolTip(_caption, hint)
        SetRange(0, 100, "")
    End Sub

    Protected Overrides Sub OnResize(ByVal e As EventArgs)
        MyBase.OnResize(e)
        _num.Left = Me.ClientSize.Width - _num.Width - 30
        _num.Top = 0
        _unit.Left = _num.Right + 4
        _unit.Top = 5
        _hint.Width = Me.ClientSize.Width - 6
        _bar.Width = Me.ClientSize.Width - 2
        _range.Width = Me.ClientSize.Width - 6
    End Sub

    Private Sub OnActivated(ByVal sender As Object, ByVal e As EventArgs)
        RaiseEvent Activated(Me, EventArgs.Empty)
    End Sub

    ' --- value ---

    Public Property Value As Double
        Get
            Return CDbl(_num.Value)
        End Get
        Set(ByVal v As Double)
            Dim clamped As Double = Clamp(v)
            If Math.Abs(clamped - CDbl(_num.Value)) < 0.0000001 Then Return
            _sync = True
            _num.Value = CDec(clamped)
            _bar.Value = ToBar(clamped)
            _sync = False
            RaiseEvent ValueChanged(Me, EventArgs.Empty)
        End Set
    End Property

    Public ReadOnly Property Minimum As Double
        Get
            Return CDbl(_num.Minimum)
        End Get
    End Property

    Public ReadOnly Property Maximum As Double
        Get
            Return CDbl(_num.Maximum)
        End Get
    End Property

    Private Function Clamp(ByVal v As Double) As Double
        If v < Minimum Then Return Minimum
        If v > Maximum Then Return Maximum
        Return v
    End Function

    Private Function ToBar(ByVal v As Double) As Integer
        Dim i As Integer = CInt(Math.Round(v * _scale))
        If i < _bar.Minimum Then i = _bar.Minimum
        If i > _bar.Maximum Then i = _bar.Maximum
        Return i
    End Function

    Private Sub OnNumChanged(ByVal sender As Object, ByVal e As EventArgs)
        If _sync Then Return
        _sync = True
        _bar.Value = ToBar(CDbl(_num.Value))
        _sync = False
        RaiseEvent ValueChanged(Me, EventArgs.Empty)
    End Sub

    Private Sub OnBarChanged(ByVal sender As Object, ByVal e As EventArgs)
        If _sync Then Return
        _sync = True
        _num.Value = CDec(_bar.Value / _scale)
        _sync = False
        RaiseEvent ValueChanged(Me, EventArgs.Empty)
    End Sub

    ' --- range ---

    ' Takes a value straight out of the model, widening the range if necessary so that it
    ' is NOT clamped on the way in.
    '
    ' The ordinary Value setter clamps, which is right for a person turning a dial and
    ' wrong for a number arriving from CATIA before the ranges have been worked out. On a
    ' cold connect the ranges are still those of the previous screen, and the plate
    ' thickness ranges depend on the table height - so reading a 23.8 mm worktop into a
    ' range that was computed from a 225 mm table collapsed it to 5 mm, silently, and the
    ' real number was gone before the correct range was ever computed. Measured 30.08.2026:
    ' worktop 23.8 -> 5.0 and shelf 15.8 -> 2.0 on every first connect.
    '
    ' The caller is expected to run ApplyLimits afterwards. That is what clamps - against
    ' the real range, derived from the real values - and it reports the change, so a model
    ' the window genuinely cannot represent still gets caught. What it no longer does is
    ' throw the value away before it has been looked at.
    Public Sub AcceptModelValue(ByVal v As Double)
        If v < Minimum OrElse v > Maximum Then
            SetRange(Math.Min(Minimum, v), Math.Max(Maximum, v), _reason)
        End If
        Value = v
    End Sub


    ' Applies a new allowed range. If the current value falls outside it, the value is
    ' pulled to the nearest end and the change is reported, so the model and the screen
    ' never disagree.
    Public Sub SetRange(ByVal min As Double, ByVal max As Double, ByVal reason As String)
        If max < min Then max = min
        Dim old As Double = If(_num.Maximum = 0 AndAlso _num.Minimum = 0, min, CDbl(_num.Value))

        _sync = True
        ' NumericUpDown and TrackBar both throw when Minimum ends up above Maximum,
        ' so widen first and narrow afterwards.
        _num.Minimum = -1000000D
        _num.Maximum = 1000000D
        _bar.Minimum = -1000000
        _bar.Maximum = 1000000

        _num.Minimum = CDec(min)
        _num.Maximum = CDec(max)
        _bar.Minimum = CInt(Math.Round(min * _scale))
        _bar.Maximum = CInt(Math.Round(max * _scale))

        ' One arrow key = one step of the shown precision (1 mm, or 0.1 mm for plates).
        ' Deriving it from the span instead gave nonsense nudges like 9 mm.
        Dim span As Double = max - min
        _bar.SmallChange = 1
        _bar.LargeChange = Math.Max(1, CInt(Math.Round(span * _scale / 20)))

        Dim v As Double = Clamp(old)
        _num.Value = CDec(v)
        _bar.Value = ToBar(v)
        _sync = False

        _reason = reason
        _rangeText = "Allowed " & DesignLimits.Fmt(min) & " - " & DesignLimits.Fmt(max) & " " & _unit.Text
        ShowRangeLine()

        If Math.Abs(v - old) > 0.0000001 Then
            RaiseEvent ValueChanged(Me, EventArgs.Empty)
        End If
    End Sub

    Private Sub ShowRangeLine()
        If Not _available Then Return
        _range.ForeColor = Theme.Muted
        _range.Text = _rangeText & If(String.IsNullOrEmpty(_reason), "", "   -   " & _reason)
        Theme.Tips.SetToolTip(_range, _range.Text)
        Theme.Tips.SetToolTip(_bar, _rangeText)
    End Sub

    ' Greys the input out with an explanation - used when the parameter has no effect in
    ' the current variant, or does not exist in the model yet.
    Public Sub SetUnavailable(ByVal note As String)
        _available = False
        _num.Enabled = False
        _bar.Enabled = False
        _caption.ForeColor = Theme.Muted
        _range.ForeColor = Theme.Warn
        _range.Text = note
        Theme.Tips.SetToolTip(_range, note)
    End Sub

    Public Sub SetAvailable()
        If _available Then Return
        _available = True
        _num.Enabled = True
        _bar.Enabled = True
        _caption.ForeColor = Theme.Ink
        ShowRangeLine()
    End Sub

    Public ReadOnly Property IsAvailable As Boolean
        Get
            Return _available
        End Get
    End Property

    Public Property Caption As String
        Get
            Return _caption.Text
        End Get
        Set(ByVal v As String)
            _caption.Text = v
        End Set
    End Property

    Public Property Hint As String
        Get
            Return _hint.Text
        End Get
        Set(ByVal v As String)
            _hint.Text = v
            Theme.Tips.SetToolTip(_caption, v)
        End Set
    End Property

End Class
