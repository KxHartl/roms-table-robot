Option Explicit On
Option Strict Off

Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports System.Windows.Forms

Public Enum SketchView
    Plan = 0       ' table from above
    Front = 1      ' table from the front
    Section = 2    ' profile cross-section
    Arm = 3        ' robot arm segments
    Path = 4       ' path P1..P6, elevation (Y across, Z up)
    PathPlan = 5   ' path P1..P6, top view (Y across, X forward)
End Enum

' Live schematic of what the numbers on the left mean.
'
' The drawing is redrawn from the current values on every change and the dimension that
' belongs to the input under the cursor is picked out in the accent colour, so "which
' number is this?" never has to be answered from the parameter name alone.
Public Class SketchPanel
    Inherits Panel

    Private _design As TableDesign
    Private _view As SketchView = SketchView.Plan
    Private _highlight As String = ""

    Public Sub New()
        Me.DoubleBuffered = True
        Me.BackColor = Theme.Card
        Me.ResizeRedraw = True
    End Sub

    Public Property Design As TableDesign
        Get
            Return _design
        End Get
        Set(ByVal v As TableDesign)
            _design = v
            Me.Invalidate()
        End Set
    End Property

    Public Property View As SketchView
        Get
            Return _view
        End Get
        Set(ByVal v As SketchView)
            If _view = v Then Return
            _view = v
            Me.Invalidate()
        End Set
    End Property

    Public Property Highlight As String
        Get
            Return _highlight
        End Get
        Set(ByVal v As String)
            If _highlight = v Then Return
            _highlight = If(v, "")
            Me.Invalidate()
        End Set
    End Property

    ' Which drawing explains a given input best.
    '
    ' A point's forward (X) and sideways (Y) sliders bring up the TOP view, its height (Z)
    ' the FRONT view - the tlocrt shows X and Y, the nacrt Z (user's call, 24.09.2026).
    ' The segments stay on the front view: the two descents are along Z, and P2-P4 is
    ' drawn in both.
    Public Shared Function ViewFor(ByVal key As String) As SketchView
        Select Case key
            Case "height", "upperPlate", "lowerPlate", "variant" : Return SketchView.Front
            Case "profile", "simplification" : Return SketchView.Section
            Case "lenJ2J3", "lenJ3J5" : Return SketchView.Arm
            Case "p1x", "p3x", "p6x", "p1y", "p3y", "p6y" : Return SketchView.PathPlan
            Case "p1", "p3", "p6", "p2p3", "p2p4", "simLeg", "simSpeed" : Return SketchView.Path
            Case Else : Return SketchView.Plan
        End Select
    End Function

    Private Function Hot(ByVal key As String) As Boolean
        Return String.Equals(_highlight, key, StringComparison.Ordinal)
    End Function

    ' One point, three sketch keys: "p3" for its height slider, "p3x" and "p3y" for forward
    ' and sideways. All must light the same dot, or switching between the two path views by
    ' hand would drop the highlight for no reason the user can see.
    Private Function HotPoint(ByVal baseKey As String) As Boolean
        Return Hot(baseKey) OrElse Hot(baseKey & "x") OrElse Hot(baseKey & "y")
    End Function

    ' --- painting ---

    Protected Overrides Sub OnPaint(ByVal e As PaintEventArgs)
        MyBase.OnPaint(e)
        Dim g As Graphics = e.Graphics
        g.SmoothingMode = SmoothingMode.AntiAlias
        g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit

        Using pen As New Pen(Theme.Border)
            g.DrawRectangle(pen, 0, 0, Me.Width - 1, Me.Height - 1)
        End Using

        Dim title As String = ""
        Dim subtitle As String = ""
        Select Case _view
            Case SketchView.Plan
                title = "Top view"
                subtitle = "Worktop seen from above. Posts and rails are dashed (hidden below the plate)."
            Case SketchView.Front
                title = "Front view"
                subtitle = "Table seen from the front, floor at the bottom."
            Case SketchView.Section
                title = "Profile cross-section"
                subtitle = "The section of one post, drawn the way it sits in the table."
            Case SketchView.Arm
                title = "Robot arm"
                subtitle = "The two segments the task lets you stretch by up to 20 %."
            Case SketchView.Path
                title = "Path, front view"
                subtitle = "P1 - P2 - P3 - P2 - P4 - P5 - P4 - P6, seen along the robot's X axis."
            Case SketchView.PathPlan
                title = "Path, top view"
                subtitle = "The same path from above: Y across as in the front view, X forward down the page."
        End Select

        Using b As New SolidBrush(Theme.Ink)
            g.DrawString(title, Theme.Heading, b, 14, 10)
        End Using
        Using b As New SolidBrush(Theme.Muted)
            g.DrawString(subtitle, Theme.Small, b, New RectangleF(14, 32, Me.Width - 28, 32))
        End Using

        DrawViewSwitch(g)

        If _design Is Nothing Then Return

        Dim area As New RectangleF(56, 66, Me.Width - 112, Me.Height - 108)
        If area.Width < 80 OrElse area.Height < 80 Then Return

        Select Case _view
            Case SketchView.Plan : DrawPlan(g, area)
            Case SketchView.Front : DrawFront(g, area)
            Case SketchView.Section : DrawSection(g, area)
            Case SketchView.Arm : DrawArm(g, area)
            Case SketchView.Path : DrawPath(g, area)
            Case SketchView.PathPlan : DrawPathPlan(g, area)
        End Select
    End Sub

    ' --- switching between the two path drawings by hand ---
    '
    ' Hovering an input already picks the drawing that explains it, and that covers the
    ' normal case. This covers the other one: having set the points, you want to look at
    ' the same path the other way round without going back to touch a slider. Two words in
    ' the corner, only on the path views - a full toolbar for two states would be more
    ' chrome than drawing.
    Private _tabTop As RectangleF = RectangleF.Empty
    Private _tabFront As RectangleF = RectangleF.Empty

    Private Function OnPathView() As Boolean
        Return _view = SketchView.Path OrElse _view = SketchView.PathPlan
    End Function

    Private Sub DrawViewSwitch(ByVal g As Graphics)
        _tabTop = RectangleF.Empty
        _tabFront = RectangleF.Empty
        If Not OnPathView() Then Return

        Dim h As Single = 22.0F
        Dim wTop As Single = g.MeasureString("Top", Theme.Small).Width + 18
        Dim wFront As Single = g.MeasureString("Front", Theme.Small).Width + 18
        Dim right As Single = Me.Width - 14
        _tabFront = New RectangleF(right - wFront, 10, wFront, h)
        _tabTop = New RectangleF(_tabFront.Left - 6 - wTop, 10, wTop, h)

        Tab_(g, _tabTop, "Top", _view = SketchView.PathPlan)
        Tab_(g, _tabFront, "Front", _view = SketchView.Path)
    End Sub

    Private Sub Tab_(ByVal g As Graphics, ByVal r As RectangleF, ByVal text As String,
                     ByVal active As Boolean)
        Using b As New SolidBrush(If(active, Theme.Accent, Theme.Card))
            g.FillRectangle(b, r)
        End Using
        Using p As New Pen(If(active, Theme.Accent, Theme.Border))
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
        End Using
        Dim sz As SizeF = g.MeasureString(text, Theme.Small)
        Using b As New SolidBrush(If(active, Theme.Card, Theme.Muted))
            g.DrawString(text, Theme.Small, b,
                         r.X + (r.Width - sz.Width) / 2, r.Y + (r.Height - sz.Height) / 2)
        End Using
    End Sub

    Protected Overrides Sub OnMouseDown(ByVal e As MouseEventArgs)
        MyBase.OnMouseDown(e)
        If Not OnPathView() Then Return
        If _tabTop.Contains(e.X, e.Y) Then Me.View = SketchView.PathPlan
        If _tabFront.Contains(e.X, e.Y) Then Me.View = SketchView.Path
    End Sub

    Protected Overrides Sub OnMouseMove(ByVal e As MouseEventArgs)
        MyBase.OnMouseMove(e)
        Dim over As Boolean = OnPathView() AndAlso
                              (_tabTop.Contains(e.X, e.Y) OrElse _tabFront.Contains(e.X, e.Y))
        Me.Cursor = If(over, Cursors.Hand, Cursors.Default)
    End Sub

    ' --- model -> pixel mapping (Y grows upwards, like the model) ---

    Private Class Mapper

        Private ReadOnly _k As Single
        Private ReadOnly _mcx As Double, _mcy As Double
        Private ReadOnly _pcx As Single, _pcy As Single

        Public Sub New(ByVal modelW As Double, ByVal modelH As Double,
                       ByVal modelCx As Double, ByVal modelCy As Double, ByVal target As RectangleF)
            If modelW <= 0 Then modelW = 1
            If modelH <= 0 Then modelH = 1
            _k = CSng(Math.Min(target.Width / modelW, target.Height / modelH))
            _mcx = modelCx
            _mcy = modelCy
            _pcx = target.Left + target.Width / 2.0F
            _pcy = target.Top + target.Height / 2.0F
        End Sub

        Public ReadOnly Property K As Single
            Get
                Return _k
            End Get
        End Property

        Public Function PX(ByVal x As Double) As Single
            Return _pcx + CSng((x - _mcx) * _k)
        End Function

        Public Function PY(ByVal y As Double) As Single
            Return _pcy - CSng((y - _mcy) * _k)
        End Function

        Public Function Len(ByVal v As Double) As Single
            Return CSng(v * _k)
        End Function

        ' Rectangle given by model corners (x0,y0)-(x1,y1).
        Public Function Box(ByVal x0 As Double, ByVal y0 As Double, ByVal x1 As Double, ByVal y1 As Double) As RectangleF
            Dim a As Single = PX(Math.Min(x0, x1))
            Dim b As Single = PY(Math.Max(y0, y1))
            Return New RectangleF(a, b, Math.Abs(PX(x1) - PX(x0)), Math.Abs(PY(y1) - PY(y0)))
        End Function

    End Class

    ' --- drawing primitives ---

    Private Sub Fill(ByVal g As Graphics, ByVal r As RectangleF, ByVal fill As Color, ByVal edge As Color)
        If r.Width < 1 Then r.Width = 1
        If r.Height < 1 Then r.Height = 1
        Using b As New SolidBrush(fill)
            g.FillRectangle(b, r)
        End Using
        Using p As New Pen(edge)
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
        End Using
    End Sub

    Private Sub Ghost(ByVal g As Graphics, ByVal r As RectangleF, ByVal edge As Color)
        Using p As New Pen(edge)
            p.DashStyle = DashStyle.Dash
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
        End Using
    End Sub

    Private Sub Note(ByVal g As Graphics, ByVal x As Single, ByVal y As Single, ByVal s As String, ByVal col As Color)
        Using b As New SolidBrush(col)
            g.DrawString(s, Theme.Small, b, x, y)
        End Using
    End Sub

    ' Same, but wrapped inside a width. The plain Note draws at a point, so a footnote that
    ' outgrows the panel simply runs off the right edge with its last words cut in half -
    ' which is what both path views did at the width this panel actually gets.
    Private Sub NoteIn(ByVal g As Graphics, ByVal x As Single, ByVal y As Single,
                       ByVal w As Single, ByVal s As String, ByVal col As Color)
        Using b As New SolidBrush(col)
            g.DrawString(s, Theme.Small, b, New RectangleF(x, y, Math.Max(w, 40.0F), 34))
        End Using
    End Sub

    Private Sub Centered(ByVal g As Graphics, ByVal cx As Single, ByVal cy As Single, ByVal s As String,
                         ByVal col As Color, ByVal bold As Boolean)
        Dim f As Font = If(bold, Theme.BodyBold, Theme.Small)
        Dim sz As SizeF = g.MeasureString(s, f)
        Using bg As New SolidBrush(Theme.Card)
            g.FillRectangle(bg, cx - sz.Width / 2 - 2, cy - sz.Height / 2 - 1, sz.Width + 4, sz.Height + 2)
        End Using
        Using b As New SolidBrush(col)
            g.DrawString(s, f, b, cx - sz.Width / 2, cy - sz.Height / 2)
        End Using
    End Sub

    Private Function DimPen(ByVal hot As Boolean) As Pen
        Dim p As New Pen(If(hot, Theme.Accent, Theme.Muted), If(hot, 1.8F, 1.0F))
        p.CustomStartCap = New AdjustableArrowCap(3.5F, 4.5F)
        p.CustomEndCap = New AdjustableArrowCap(3.5F, 4.5F)
        Return p
    End Function

    ' Horizontal dimension between two pixel columns, drawn at pixel row y.
    Private Sub DimH(ByVal g As Graphics, ByVal xa As Single, ByVal xb As Single, ByVal y As Single,
                     ByVal label As String, ByVal hot As Boolean)
        Using p As Pen = DimPen(hot)
            g.DrawLine(p, xa, y, xb, y)
        End Using
        Using p As New Pen(Color.FromArgb(90, Theme.Muted))
            g.DrawLine(p, xa, y - 6, xa, y + 6)
            g.DrawLine(p, xb, y - 6, xb, y + 6)
        End Using
        Centered(g, (xa + xb) / 2, y, label, If(hot, Theme.Accent, Theme.Muted), hot)
    End Sub

    ' Vertical dimension between two pixel rows, drawn at pixel column x.
    Private Sub DimV(ByVal g As Graphics, ByVal ya As Single, ByVal yb As Single, ByVal x As Single,
                     ByVal label As String, ByVal hot As Boolean)
        Using p As Pen = DimPen(hot)
            g.DrawLine(p, x, ya, x, yb)
        End Using
        Using p As New Pen(Color.FromArgb(90, Theme.Muted))
            g.DrawLine(p, x - 6, ya, x + 6, ya)
            g.DrawLine(p, x - 6, yb, x + 6, yb)
        End Using
        Centered(g, x, (ya + yb) / 2, label, If(hot, Theme.Accent, Theme.Muted), hot)
    End Sub

    Private Function Mm(ByVal v As Double) As String
        Return DesignLimits.Fmt(v) & " mm"
    End Function

    ' --- top view ---

    Private Sub DrawPlan(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim m As New Mapper(d.Length * 1.22, d.Width * 1.35, 0, 0, area)

        ' worktop
        Fill(g, m.Box(-d.Length / 2, -d.Width / 2, d.Length / 2, d.Width / 2), Theme.PlateFill, Theme.SteelDark)

        ' posts and rails, hidden under the plate
        Dim px As Double = d.PostCenterX, py As Double = d.PostCenterY
        Dim ol As Double = d.OffsetLength, ow As Double = d.OffsetWidth
        Ghost(g, m.Box(-px + ol, -py - ow, px - ol, -py + ow), Theme.SteelDark)
        Ghost(g, m.Box(-px + ol, py - ow, px - ol, py + ow), Theme.SteelDark)
        Ghost(g, m.Box(-px - ol, -py + ow, -px + ol, py - ow), Theme.SteelDark)
        Ghost(g, m.Box(px - ol, -py + ow, px + ol, py - ow), Theme.SteelDark)

        For Each sx As Integer In New Integer() {-1, 1}
            For Each sy As Integer In New Integer() {-1, 1}
                Fill(g, m.Box(sx * px - ol, sy * py - ow, sx * px + ol, sy * py + ow),
                     Theme.Steel, Theme.SteelDark)
            Next
        Next

        ' robot base: bolt raster plus a dashed footprint around it
        Dim hx As Double() = d.MountHolesX()
        Dim hy As Double() = d.MountHolesY()
        Dim x0 As Double = Math.Min(hx(0), hx(1)), x1 As Double = Math.Max(hx(0), hx(1))
        Dim y0 As Double = Math.Min(hy(0), hy(1)), y1 As Double = Math.Max(hy(0), hy(1))
        ' Dashed outline = the base body, the bolt raster grown by its overhang. This is
        ' what must not hang over the plate edge, not the bolt holes themselves.
        Dim pad As Double = d.BaseOverhang
        Dim hotBase As Boolean = Hot("baseOverhang")
        Using p As New Pen(If(hotBase, Theme.Accent, Theme.Accent), If(hotBase, 2.4F, 1.4F))
            p.DashStyle = DashStyle.Dash
            Dim r As RectangleF = m.Box(x0 - pad, y0 - pad, x1 + pad, y1 + pad)
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
            If hotBase Then
                Using b As New SolidBrush(Color.FromArgb(28, Theme.Accent))
                    g.FillRectangle(b, r)
                End Using
            End If
        End Using

        Dim rr As Single = Math.Max(2.5F, m.Len(d.MountHoleDiameter / 2))
        For Each cx As Double In hx
            For Each cy As Double In hy
                Using b As New SolidBrush(Theme.Card)
                    g.FillEllipse(b, m.PX(cx) - rr, m.PY(cy) - rr, rr * 2, rr * 2)
                End Using
                Using p As New Pen(Theme.Accent, 1.4F)
                    g.DrawEllipse(p, m.PX(cx) - rr, m.PY(cy) - rr, rr * 2, rr * 2)
                End Using
            Next
        Next
        ' Inside the raster: at the edge of the plate the label would run off the drawing.
        Centered(g, m.PX((x0 + x1) / 2), m.PY((y0 + y1) / 2), "robot base", Theme.Accent, False)

        ' overall dimensions
        DimH(g, m.PX(-d.Length / 2), m.PX(d.Length / 2), m.PY(-d.Width / 2) + 24,
             "length  " & Mm(d.Length), Hot("length"))
        DimV(g, m.PY(-d.Width / 2), m.PY(d.Width / 2), m.PX(-d.Length / 2) - 26,
             "width  " & Mm(d.Width), Hot("width"))

        ' the offsets that decide where the robot stands
        If d.UsesOffsetX Then
            Dim yLine As Single = m.PY(y0 - pad) + 16
            DimH(g, m.PX(d.AxisOffsetX), m.PX(d.Length / 2), yLine,
                 "from edge  " & Mm(d.UserOffsetX), Hot("offsetX"))
        End If
        If d.UsesOffsetY Then
            Dim xLine As Single = m.PX(x0 - pad) - 16
            DimV(g, m.PY(d.AxisOffsetY), m.PY(d.Width / 2), xLine,
                 "from edge  " & Mm(d.UserOffsetY), Hot("offsetY"))
        End If

        ' axis rose
        Dim ax As Single = area.Right - 4, ay As Single = area.Bottom + 6
        Using p As New Pen(Theme.Muted, 1.0F)
            p.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(p, ax - 34, ay, ax - 6, ay)
            g.DrawLine(p, ax - 34, ay, ax - 34, ay - 26)
        End Using
        Note(g, ax - 6, ay - 8, "X", Theme.Muted)
        Note(g, ax - 44, ay - 30, "Y", Theme.Muted)
    End Sub

    ' --- front view ---

    Private Sub DrawFront(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim m As New Mapper(d.Length * 1.25, d.Height * 1.3, 0, d.Height / 2, area)

        Dim px As Double = d.PostCenterX, ol As Double = d.OffsetLength
        Dim fh As Double = d.Constants.FeetHeight
        Dim topOfPost As Double = fh + d.HeightVerticalProfile

        ' floor
        Using p As New Pen(Theme.SteelDark, 1.6F)
            g.DrawLine(p, m.PX(-d.Length / 2) - 22, m.PY(0), m.PX(d.Length / 2) + 22, m.PY(0))
        End Using
        Using p As New Pen(Color.FromArgb(120, Theme.Muted))
            Dim x As Single = m.PX(-d.Length / 2) - 20
            While x < m.PX(d.Length / 2) + 22
                g.DrawLine(p, x, m.PY(0), x - 7, m.PY(0) + 7)
                x += 11
            End While
        End Using

        ' feet and posts
        For Each s As Integer In New Integer() {-1, 1}
            Fill(g, m.Box(s * px - ol, 0, s * px + ol, fh), Theme.SteelDark, Theme.SteelDark)
            Fill(g, m.Box(s * px - ol, fh, s * px + ol, topOfPost), Theme.Steel, Theme.SteelDark)
        Next

        ' rails between the posts
        Dim railL As Double = -(px - ol), railR As Double = px - ol
        Dim upperRailZ As Double = d.Height - d.UpperPlateThickness - ol
        Fill(g, m.Box(railL, upperRailZ - ol, railR, upperRailZ + ol), Theme.Steel, Theme.SteelDark)
        Fill(g, m.Box(railL, d.HeightLowerHorizontal - ol, railR, d.HeightLowerHorizontal + ol),
             Theme.Steel, Theme.SteelDark)

        ' brackets: this is the whole difference between the two shelf variants
        Dim gs As Double = d.GussetSize
        Dim bracketHot As Boolean = Hot("variant")
        Dim bracketCol As Color = If(bracketHot, Theme.Accent, Theme.SteelDark)
        For Each s As Integer In New Integer() {-1, 1}
            ' under the upper rail, always
            TriangleBracket(g, m, s, px - ol, upperRailZ - ol, gs, False, Theme.SteelDark)
            ' at the lower rail, above in V2 and below in V1
            TriangleBracket(g, m, s, px - ol,
                            If(d.GussetAbove, d.HeightLowerHorizontal + ol, d.HeightLowerHorizontal - ol),
                            gs, d.GussetAbove, bracketCol)
        Next

        ' plates
        Fill(g, m.Box(-d.Length / 2, d.Height - d.UpperPlateThickness, d.Length / 2, d.Height),
             Theme.PlateFill, Theme.SteelDark)
        Fill(g, m.Box(-d.LowerPlateLength / 2, d.LowerPlateZ, d.LowerPlateLength / 2,
                      d.LowerPlateZ + d.LowerPlateThickness), Theme.PlateFill, Theme.SteelDark)

        ' dimensions
        Dim rightX As Single = m.PX(d.Length / 2) + 30
        DimV(g, m.PY(0), m.PY(d.Height), rightX, "height  " & Mm(d.Height), Hot("height"))

        Dim leftX As Single = m.PX(-d.Length / 2) - 28
        DimV(g, m.PY(0), m.PY(fh), leftX, "feet " & Mm(fh), False)
        DimV(g, m.PY(d.Height - d.UpperPlateThickness), m.PY(d.Height), leftX,
             "worktop " & Mm(d.UpperPlateThickness), Hot("upperPlate"))
        DimV(g, m.PY(d.LowerPlateZ), m.PY(d.LowerPlateZ + d.LowerPlateThickness), leftX,
             "shelf " & Mm(d.LowerPlateThickness), Hot("lowerPlate"))

        Centered(g, m.PX(0), m.PY(d.LowerPlateZ) - 12,
                 "shelf  " & Mm(d.LowerPlateLength) & " x " & Mm(d.LowerPlateWidth),
                 If(Hot("variant"), Theme.Accent, Theme.Muted), Hot("variant"))
        Note(g, area.Left, area.Bottom + 4,
             If(d.GussetAbove, "V2: brackets sit on top of the lower rails, shelf flush with the post ends.",
                               "V1: brackets hang under the lower rails, so the whole lower level climbs by " & Mm(gs) & "."),
             If(bracketHot, Theme.Accent, Theme.Muted))
    End Sub

    ' Right-angled bracket drawn at the post, either above or below the rail.
    Private Sub TriangleBracket(ByVal g As Graphics, ByVal m As Mapper, ByVal side As Integer,
                                ByVal postInnerX As Double, ByVal z As Double, ByVal size As Double,
                                ByVal above As Boolean, ByVal col As Color)
        Dim x0 As Double = side * postInnerX
        Dim x1 As Double = x0 - side * size
        Dim z1 As Double = If(above, z + size, z - size)
        Dim pts As PointF() = {
            New PointF(m.PX(x0), m.PY(z)),
            New PointF(m.PX(x1), m.PY(z)),
            New PointF(m.PX(x0), m.PY(z1))
        }
        Using b As New SolidBrush(Color.FromArgb(70, col))
            g.FillPolygon(b, pts)
        End Using
        Using p As New Pen(col, 1.2F)
            g.DrawPolygon(p, pts)
        End Using
    End Sub

    ' --- profile cross-section ---

    Private Sub DrawSection(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim pr As ProfileInfo = d.Profile
        Dim sx As Double = pr.SizeAlongLength   ' along X
        Dim sy As Double = pr.SizeAlongWidth    ' along Y
        Dim m As New Mapper(sx * 1.7, sy * 1.9, 0, 0, area)

        Dim outer As RectangleF = m.Box(-sx / 2, -sy / 2, sx / 2, sy / 2)
        Fill(g, outer, Theme.PlateFill, Theme.SteelDark)

        If d.Simplified Then
            Centered(g, outer.Left + outer.Width / 2, outer.Top + outer.Height / 2,
                     "solid block", Theme.SteelDark, True)
        Else
            ' one 45 mm module per bore; the 90 mm faces really do carry two bolt rows
            Dim nx As Integer = Math.Max(1, CInt(Math.Round(sx / 45.0)))
            Dim ny As Integer = Math.Max(1, CInt(Math.Round(sy / 45.0)))
            For ix As Integer = 0 To nx - 1
                For iy As Integer = 0 To ny - 1
                    Dim cx As Double = -sx / 2 + sx / nx * (ix + 0.5)
                    Dim cy As Double = -sy / 2 + sy / ny * (iy + 0.5)
                    Dim r As Single = m.Len(4.0)
                    Using p As New Pen(Theme.SteelDark)
                        g.DrawEllipse(p, m.PX(cx) - r, m.PY(cy) - r, r * 2, r * 2)
                    End Using
                Next
            Next
            ' slot per module on every face
            For ix As Integer = 0 To nx - 1
                Dim cx As Double = -sx / 2 + sx / nx * (ix + 0.5)
                Slot(g, m, cx, sy / 2, 12, 7, False)
                Slot(g, m, cx, -sy / 2, 12, -7, False)
            Next
            For iy As Integer = 0 To ny - 1
                Dim cy As Double = -sy / 2 + sy / ny * (iy + 0.5)
                Slot(g, m, -sx / 2, cy, 12, 7, True)
                Slot(g, m, sx / 2, cy, 12, -7, True)
            Next
        End If

        Dim hotP As Boolean = Hot("profile") OrElse Hot("simplification")
        DimH(g, outer.Left, outer.Right, outer.Bottom + 26,
             DesignLimits.Fmt(sx) & " mm along the table length (X)", hotP)
        DimV(g, outer.Bottom, outer.Top, outer.Left - 26,
             DesignLimits.Fmt(sy) & " mm across the width (Y)", hotP)

        ' Two bracket sizes, and on a non-square section they are not the same one: the
        ' post joints follow the width, the rail-to-rail corners the height of a rail.
        Dim brackets As String = Mm(pr.GussetSize)
        If Math.Abs(pr.GussetSizeCorner - pr.GussetSize) > 0.001 Then
            brackets = Mm(pr.GussetSize) & " on the posts, " &
                       Mm(pr.GussetSizeCorner) & " in the corners"
        End If
        Note(g, area.Left, area.Bottom - 30,
             "Brackets that go with this section: " & brackets &
             ".  The upright posts stand on the " & DesignLimits.Fmt(sx) & " mm side.",
             Theme.Muted)
        Note(g, area.Left, area.Bottom - 14,
             If(d.Simplified, "Simplified: slots and bores are dropped, the outer envelope stays identical.",
                              "Full section: slots and bores are modelled."),
             If(Hot("simplification"), Theme.Accent, Theme.Muted))
    End Sub

    ' T-slot notch cut into one face of the profile.
    Private Sub Slot(ByVal g As Graphics, ByVal m As Mapper, ByVal cx As Double, ByVal cy As Double,
                     ByVal opening As Double, ByVal depth As Double, ByVal vertical As Boolean)
        Dim r As RectangleF
        If vertical Then
            r = m.Box(cx, cy - opening / 2, cx + depth, cy + opening / 2)
        Else
            r = m.Box(cx - opening / 2, cy, cx + opening / 2, cy - depth)
        End If
        Using b As New SolidBrush(Theme.Card)
            g.FillRectangle(b, r)
        End Using
        Using p As New Pen(Theme.SteelDark)
            g.DrawRectangle(p, r.X, r.Y, r.Width, r.Height)
        End Using
    End Sub

    ' --- robot arm ---

    Private Sub DrawArm(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim baseH As Double = 330.0
        Dim a As Double = d.LenJ2J3, b As Double = d.LenJ3J5
        Dim total As Double = baseH + a + b
        Dim m As New Mapper(total * 1.5, total * 1.25, total * 0.15, total * 0.45, area)

        ' base
        Dim j2 As New PointF(m.PX(0), m.PY(baseH))
        Fill(g, m.Box(-90, 0, 90, 60), Theme.Steel, Theme.SteelDark)
        Fill(g, m.Box(-55, 60, 55, baseH), Theme.Steel, Theme.SteelDark)
        Note(g, m.PX(-90), m.PY(0) + 6, "table top", Theme.Muted)

        ' upper arm J2 -> J3 leans back, forearm J3 -> J5 runs forward
        Dim ang As Double = 72.0 * Math.PI / 180.0
        Dim j3 As New PointF(m.PX(-a * Math.Cos(ang)), m.PY(baseH + a * Math.Sin(ang)))
        Dim j5 As New PointF(m.PX(-a * Math.Cos(ang) + b), m.PY(baseH + a * Math.Sin(ang)))

        Using p As New Pen(If(Hot("lenJ2J3"), Theme.Accent, Theme.SteelDark), 14.0F)
            p.StartCap = LineCap.Round
            p.EndCap = LineCap.Round
            g.DrawLine(p, j2, j3)
        End Using
        Using p As New Pen(If(Hot("lenJ3J5"), Theme.Accent, Theme.SteelDark), 12.0F)
            p.StartCap = LineCap.Round
            p.EndCap = LineCap.Round
            g.DrawLine(p, j3, j5)
        End Using

        For Each jp As PointF In New PointF() {j2, j3, j5}
            Using bb As New SolidBrush(Theme.Card)
                g.FillEllipse(bb, jp.X - 6, jp.Y - 6, 12, 12)
            End Using
            Using p As New Pen(Theme.Ink, 1.5F)
                g.DrawEllipse(p, jp.X - 6, jp.Y - 6, 12, 12)
            End Using
        Next
        Note(g, j2.X + 10, j2.Y + 4, "J2", Theme.Ink)
        Note(g, j3.X - 6, j3.Y - 22, "J3", Theme.Ink)
        Note(g, j5.X + 8, j5.Y - 6, "J5", Theme.Ink)

        Dim nomA As Double = d.Constants.NominalLenJ2J3
        Dim nomB As Double = d.Constants.NominalLenJ3J5
        Centered(g, (j2.X + j3.X) / 2 - 46, (j2.Y + j3.Y) / 2,
                 "J2-J3  " & Mm(a) & "  (+" & DesignLimits.Fmt((a / nomA - 1) * 100) & " %)",
                 If(Hot("lenJ2J3"), Theme.Accent, Theme.Muted), Hot("lenJ2J3"))
        Centered(g, (j3.X + j5.X) / 2, j3.Y - 24,
                 "J3-J5  " & Mm(b) & "  (+" & DesignLimits.Fmt((b / nomB - 1) * 100) & " %)",
                 If(Hot("lenJ3J5"), Theme.Accent, Theme.Muted), Hot("lenJ3J5"))

        Note(g, area.Left, area.Bottom - 14,
             "Nominal from RoboDK: J2-J3 " & Mm(nomA) & ", J3-J5 " & Mm(nomB) & ". The task allows up to +20 %.",
             Theme.Muted)
    End Sub

    ' --- simulation path ---

    ' A real elevation of the path, looking along the robot's X axis: Y across, Z up.
    '
    ' It used to be a schematic - P2P4/P2P3/P4P5 and a hard-coded 140 mm "lift", with P1 and
    ' P6 drawn at invented positions. That was survivable while P1 was the only free point,
    ' but once P1, P3 and P6 are all typed in, a picture that ignores them is a picture that
    ' lies. Everything below comes from RobotKinematics.Point, the same function the arm and
    ' the checks use.
    Private Sub DrawPath(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim inp As RobotKinematics.PathInputs = d.ToPathInputs()

        Dim names As String() = {"P1", "P2", "P3", "P4", "P5", "P6"}
        Dim yv(5) As Double
        Dim zv(5) As Double
        Dim yMin As Double = Double.MaxValue, yMax As Double = Double.MinValue
        Dim zMin As Double = Double.MaxValue, zMax As Double = Double.MinValue

        For i As Integer = 0 To 5
            Dim c As Double() = RobotKinematics.Point(inp, names(i))
            yv(i) = c(1)
            zv(i) = c(2)
            yMin = Math.Min(yMin, yv(i)) : yMax = Math.Max(yMax, yv(i))
            zMin = Math.Min(zMin, zv(i)) : zMax = Math.Max(zMax, zv(i))
        Next

        ' Keep a margin so the labels have somewhere to sit, and never divide by zero when
        ' the six points happen to line up.
        Dim spanY As Double = Math.Max(yMax - yMin, 1.0) * 1.35
        Dim spanZ As Double = Math.Max(zMax - zMin, 1.0) * 1.45
        ' Short of the bottom edge, so the P2-P4 dimension under the stations and the
        ' footnote do not land on each other - they did.
        Dim plot As New RectangleF(area.X, area.Y, area.Width, area.Height - 34)
        Dim m As New Mapper(spanY, spanZ, (yMin + yMax) / 2.0, (zMin + zMax) / 2.0, plot)

        Dim pt(5) As PointF
        For i As Integer = 0 To 5
            pt(i) = New PointF(m.PX(yv(i)), m.PY(zv(i)))
        Next
        Dim p1 = pt(0), p2 = pt(1), p3 = pt(2), p4 = pt(3), p5 = pt(4), p6 = pt(5)

        ' Straight legs solid, joint legs dashed - the same distinction the simulation
        ' makes, so the picture says which moves are MoveL and which are MoveJ.
        Using pj As New Pen(Theme.Muted, 2.0F)
            pj.DashStyle = Drawing2D.DashStyle.Dash
            pj.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(pj, p1, p2)
            g.DrawLine(pj, p4, p6)
        End Using

        Using p As New Pen(Theme.Steel, 2.0F)
            p.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(p, p2, p3)
            g.DrawLine(p, p3, p2)
            g.DrawLine(p, p2, p4)
            g.DrawLine(p, p4, p5)
            g.DrawLine(p, p5, p4)
        End Using

        Dim keys As String() = {"p1", "p2p3", "p3", "p2p4", "p3", "p6"}
        For i As Integer = 0 To 5
            Dim live As Boolean = HotPoint(keys(i))
            Using bb As New SolidBrush(If(live, Theme.Accent, Theme.Card))
                g.FillEllipse(bb, pt(i).X - 5, pt(i).Y - 5, 10, 10)
            End Using
            Using p As New Pen(If(live, Theme.Accent, Theme.Ink), 1.5F)
                g.DrawEllipse(p, pt(i).X - 5, pt(i).Y - 5, 10, 10)
            End Using
            Note(g, pt(i).X + 8, pt(i).Y - 16, names(i), Theme.Ink)
        Next

        DimV(g, p2.Y, p3.Y, Math.Min(p2.X, p3.X) - 52, "P2-P3  " & Mm(d.P2P3), Hot("p2p3"))
        DimH(g, p2.X, p4.X, Math.Max(p2.Y, p4.Y) + 44, "P2-P4  " & Mm(d.P2P4), Hot("p2p4"))

        ShowPoint(g, p1, "P1", d.P1X, d.P1Y, d.P1Z, HotPoint("p1"))
        ShowPoint(g, p3, "P3", d.P3X, d.P3Y, d.P3Z, HotPoint("p3"))
        ShowPoint(g, p6, "P6", d.P6X, d.P6Y, d.P6Z, HotPoint("p6"))

        NoteIn(g, area.Left, area.Bottom - 28, area.Width - 56,
               "Along the robot's X axis - forward is edge-on, which is what the top " &
               "view shows. Solid = MoveL, dashed = MoveJ.", Theme.Muted)
    End Sub


    ' --- the path from above ---
    '
    ' The elevation looks straight down the robot's X axis, which is exactly the direction
    ' P1_x, P3_x and P6_x move in: three of the nine coordinates the user types are
    ' invisible in it. This is the other half of the drawing.
    '
    ' It is also the only view in which REACH is a shape rather than a number. The arm
    ' sweeps an annulus about the J1 axis - outer radius the full reach, inner radius the
    ' fold the elbow cannot come inside - and whether a station is in it is a question about
    ' sqrt(x^2 + y^2), which the eye answers instantly on a circle and not at all on a pair
    ' of sliders. The two circles are the same numbers DesignLimits puts on the forward
    ' slider, so the picture and the limit cannot drift apart.
    '
    ' AXES: Y to the RIGHT, X DOWN the page. First-angle projection, which is what a
    ' Croatian engineering drawing uses: the tlocrt sits under the nacrt and shares its
    ' horizontal axis with it. Two consequences worth having - a point keeps its horizontal
    ' place when you switch Top/Front, and the robot is at the top of the sheet reaching
    ' down it, which is how the pair reads as one drawing rather than two.
    '
    ' It took three tries (user's call each time): X up / Y right, then X right / Y up,
    ' then this. The mapper only ever gets Y horizontally and -X vertically - PY grows
    ' upwards, so negating X is what puts forward at the bottom.
    Private Sub DrawPathPlan(ByVal g As Graphics, ByVal area As RectangleF)
        Dim d As TableDesign = _design
        Dim inp As RobotKinematics.PathInputs = d.ToPathInputs()

        ' In plan P2 sits on P3 and P5 on P4 - the descents are along Z. So four distinct
        ' places, not six, and saying which two are stacked is part of what the view says.
        Dim c1 As Double() = RobotKinematics.Point(inp, "P1")
        Dim c3 As Double() = RobotKinematics.Point(inp, "P3")
        Dim c4 As Double() = RobotKinematics.Point(inp, "P4")
        Dim c6 As Double() = RobotKinematics.Point(inp, "P6")

        Dim xs As Double() = {0.0, c1(0), c3(0), c4(0), c6(0)}
        Dim ys As Double() = {0.0, c1(1), c3(1), c4(1), c6(1)}
        Dim xMin As Double = Double.MaxValue, xMax As Double = Double.MinValue
        Dim yMin As Double = Double.MaxValue, yMax As Double = Double.MinValue
        For i As Integer = 0 To xs.Length - 1
            xMin = Math.Min(xMin, xs(i)) : xMax = Math.Max(xMax, xs(i))
            yMin = Math.Min(yMin, ys(i)) : yMax = Math.Max(yMax, ys(i))
        Next

        ' The robot itself is always in frame - it is the origin every dimension is taken
        ' from - and the margin leaves room for the labels.
        Dim spanX As Double = Math.Max(xMax - xMin, 1.0) * 1.5
        Dim spanY As Double = Math.Max(yMax - yMin, 1.0) * 1.45
        ' A shorter box than the other views get: the footnote is two lines here, and
        ' without the gap it prints over the drawing.
        Dim plot As New RectangleF(area.X, area.Y, area.Width, area.Height - 34)
        ' PX maps Y (sideways, to the right). PY maps -X, because the mapper's vertical
        ' axis grows upwards and forward has to go DOWN: every X below is passed negated.
        Dim m As New Mapper(spanY, spanX, (yMin + yMax) / 2.0, -(xMin + xMax) / 2.0, plot)

        Dim reach As Double = RobotKinematics.Reach(inp)
        Dim inner As Double = RobotKinematics.RB_A1 +
                              Math.Abs(inp.LenJ2J3 - RobotKinematics.L3(inp))

        ' The circles run well outside the six points, so they are clipped to the drawing
        ' area instead of being allowed to paint over the title and the notes.
        Dim oldClip As Region = g.Clip
        g.SetClip(plot)
        Circle(g, m, reach, Theme.Accent, 1.4F, True)
        Circle(g, m, inner, Theme.Muted, 1.0F, True)
        g.Clip = oldClip

        ' robot: the J1 axis, and the direction its own X points in
        Dim o As New PointF(m.PX(0), m.PY(0))
        Using p As New Pen(Theme.SteelDark, 1.4F)
            g.DrawEllipse(p, o.X - 7, o.Y - 7, 14, 14)
            g.DrawLine(p, o.X - 12, o.Y, o.X + 12, o.Y)
            g.DrawLine(p, o.X, o.Y - 12, o.X, o.Y + 12)
        End Using
        Note(g, o.X + 11, o.Y + 2, "J1", Theme.Muted)

        Dim q1 As New PointF(m.PX(c1(1)), m.PY(-c1(0)))
        Dim q3 As New PointF(m.PX(c3(1)), m.PY(-c3(0)))
        Dim q4 As New PointF(m.PX(c4(1)), m.PY(-c4(0)))
        Dim q6 As New PointF(m.PX(c6(1)), m.PY(-c6(0)))

        ' Same convention as the elevation: solid is MoveL, dashed is MoveJ.
        Using pj As New Pen(Theme.Muted, 2.0F)
            pj.DashStyle = DashStyle.Dash
            pj.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(pj, q1, q3)
            g.DrawLine(pj, q4, q6)
        End Using
        Using p As New Pen(Theme.Steel, 2.0F)
            p.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(p, q3, q4)
        End Using

        ' HOW FAR OUT the far station is, drawn as the radius it actually is.
        '
        ' The reach circle is honest but usually invisible: at 704 mm it falls outside a
        ' drawing scaled to a path that reaches 420, so the footnote promised a circle the
        ' user could not see. This line is the same fact where it can be read - from the J1
        ' axis to the station that is furthest out, labelled with both numbers.
        Dim rP3 As Double = Math.Sqrt(c3(0) * c3(0) + c3(1) * c3(1))
        Dim rP4 As Double = Math.Sqrt(c4(0) * c4(0) + c4(1) * c4(1))
        Dim far As PointF = If(rP4 > rP3, q4, q3)
        Dim rFar As Double = Math.Max(rP3, rP4)
        Using p As New Pen(Color.FromArgb(120, Theme.Accent), 1.0F)
            p.DashStyle = DashStyle.Dot
            g.DrawLine(p, o, far)
        End Using
        ' No floating label on it. With forward running down the page this line lies along
        ' the same middle band as the forward dimension, and two labels were fighting over
        ' one strip of drawing. The number goes in the footnote instead, where it also gets
        ' to stand next to the reach it should be compared with.

        ' The two stations carry a second point each - the lift and the descent are along Z
        ' and collapse to nothing here, which is worth saying rather than leaving as a
        ' silent disagreement with the other view.
        Station(g, q3, "P3 / P2", HotPoint("p3"))
        Station(g, q4, "P5 / P4", Hot("p2p4"))
        ' Label on the outward side of each dot - away from the centre line, where the
        ' dimensions and the other point's name are.
        Dot(g, q1, "P1", HotPoint("p1"), c1(1) < 0)
        Dot(g, q6, "P6", HotPoint("p6"), c6(1) < 0)

        ' Forward from the robot's own axis, and across it: the two the elevation cannot
        ' dimension. P2P4 is here too, and in this view it is plainly a straight move.
        '
        ' Forward is vertical now and takes the left margin; sideways sits in the gap
        ' between the robot and the stations, and P2-P4 just under the two stations - three
        ' sides of the drawing, so no two labels share a spot.
        DimV(g, m.PY(0), m.PY(-c3(0)), Math.Min(o.X, q3.X) - 44,
             "P3 forward  " & Mm(d.P3X), Hot("p3x"))
        DimH(g, m.PX(0), m.PX(c3(1)), m.PY(0) + 30,
             "P3 sideways  " & Mm(d.P3Y), HotPoint("p3"))
        DimH(g, q3.X, q4.X, Math.Max(q3.Y, q4.Y) + 28, "P2-P4  " & Mm(d.P2P4), Hot("p2p4"))

        ShowPoint(g, q1, "P1", d.P1X, d.P1Y, d.P1Z, HotPoint("p1"))
        ShowPoint(g, q3, "P3", d.P3X, d.P3Y, d.P3Z, HotPoint("p3"))
        ShowPoint(g, q6, "P6", d.P6X, d.P6Y, d.P6Z, HotPoint("p6"))

        ' axis rose: Y right, X DOWN - so its corner is at the top, not at the bottom like
        ' the roses on the table views.
        Dim ax As Single = area.Right - 4, ay As Single = area.Bottom + 6
        Dim cx As Single = ax - 34, cy As Single = ay - 26
        Using p As New Pen(Theme.Muted, 1.0F)
            p.CustomEndCap = New AdjustableArrowCap(3, 4)
            g.DrawLine(p, cx, cy, ax - 6, cy)
            g.DrawLine(p, cx, cy, cx, ay)
        End Using
        Note(g, ax - 6, cy - 8, "Y", Theme.Muted)
        Note(g, cx - 16, ay - 10, "X", Theme.Muted)

        NoteIn(g, area.Left, area.Bottom - 28, area.Width - 56,
               "Dotted: far station at " & Mm(rFar) & " of " & Mm(reach) &
               " reach, fold " & Mm(inner) & ". P2 sits over P3, P5 over P4.", Theme.Muted)
    End Sub

    ' Circle of model radius r about the robot's J1 axis.
    Private Sub Circle(ByVal g As Graphics, ByVal m As Mapper, ByVal r As Double,
                       ByVal col As Color, ByVal width As Single, ByVal dashed As Boolean)
        Dim rp As Single = m.Len(r)
        If rp <= 0 Then Return
        Using p As New Pen(col, width)
            If dashed Then p.DashStyle = DashStyle.Dash
            g.DrawEllipse(p, m.PX(0) - rp, m.PY(0) - rp, rp * 2, rp * 2)
        End Using
    End Sub

    ' A station: two path points on one spot, drawn as a square so it does not read as the
    ' same kind of mark as P1 and P6.
    Private Sub Station(ByVal g As Graphics, ByVal at As PointF, ByVal label As String,
                        ByVal live As Boolean)
        Dim col As Color = If(live, Theme.Accent, Theme.Ink)
        Using b As New SolidBrush(If(live, Theme.Accent, Theme.Card))
            g.FillRectangle(b, at.X - 5, at.Y - 5, 10, 10)
        End Using
        Using p As New Pen(col, 1.5F)
            g.DrawRectangle(p, at.X - 5, at.Y - 5, 10, 10)
        End Using
        Note(g, at.X + 9, at.Y - 16, label, Theme.Ink)
    End Sub

    Private Sub Dot(ByVal g As Graphics, ByVal at As PointF, ByVal label As String,
                    ByVal live As Boolean, ByVal labelLeft As Boolean)
        Using b As New SolidBrush(If(live, Theme.Accent, Theme.Card))
            g.FillEllipse(b, at.X - 5, at.Y - 5, 10, 10)
        End Using
        Using p As New Pen(If(live, Theme.Accent, Theme.Ink), 1.5F)
            g.DrawEllipse(p, at.X - 5, at.Y - 5, 10, 10)
        End Using
        Note(g, If(labelLeft, at.X - 26, at.X + 8), at.Y - 16, label, Theme.Ink)
    End Sub

    Private Sub ShowPoint(ByVal g As Graphics, ByVal at As PointF, ByVal name As String,
                          ByVal x As Double, ByVal y As Double, ByVal z As Double,
                          ByVal live As Boolean)
        If Not live Then Return
        Centered(g, at.X, at.Y - 24,
                 name & " = (" & DesignLimits.Fmt(x) & ", " & DesignLimits.Fmt(y) & ", " &
                 DesignLimits.Fmt(z) & ")", Theme.Accent, True)
    End Sub

End Class
