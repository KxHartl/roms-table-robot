Option Explicit On
Option Strict Off

' Allowed range of one input, plus the sentence that explains the binding limit.
Public Class ValueRange

    Public ReadOnly Minimum As Double
    Public ReadOnly Maximum As Double
    Public ReadOnly Reason As String

    Public Sub New(ByVal min_ As Double, ByVal max_ As Double, ByVal reason_ As String)
        Minimum = min_
        Maximum = If(max_ < min_, min_, max_)
        Reason = reason_
    End Sub

End Class

Public Enum CheckLevel
    Info = 0
    Warning = 1
    Problem = 2
End Enum

Public Class DesignCheck

    Public ReadOnly Level As CheckLevel
    Public ReadOnly Text As String

    Public Sub New(ByVal level_ As CheckLevel, ByVal text_ As String)
        Level = level_
        Text = text_
    End Sub

End Class

' Every limit the UI enforces, derived from the same formulas the model uses.
'
' Design decision: a limit is enforced when it would break an update, and only warned
' about when it is merely unusual. Limits are computed for BOTH bracket variants where
' they differ, so switching V1 <-> V2 can never turn a valid table into a broken one.
Public Module DesignLimits

    Public Const MinClearPlate As Double = 40.0     ' narrowest usable lower plate
    Public Const MinRailHeadroom As Double = 20.0   ' spare room around the brackets
    Public Const MaxTableSize As Double = 2000.0
    Public Const MaxTableHeight As Double = 2000.0

    ' --- table size ---

    ' How far the OUTER bolt has to stay from the plate edge: far enough for the hole not
    ' to break the edge, and far enough for the base body not to hang over it.
    Public Function KeepFromEdge(ByVal d As TableDesign) As Double
        Return Math.Max(d.Constants.HoleEdgeDistance, d.BaseOverhang)
    End Function

    ' Reserve that keeps user_offset_x usable even though the master's hole sketch
    ' currently applies it across the WIDTH instead of along the length (90 degree
    ' rotation, see PARAMETERS.md -> R_RobotMountOffset). Without it a value that is
    ' perfectly safe on a 600 mm length lands 70 mm outside a 400 mm width.
    ' Goes to zero the moment the sketch is corrected.
    Private Function CrossReserveX(ByVal d As TableDesign) As Double
        If d.Constants.MountSketchFixed Then Return 0
        Return Math.Max(0, (d.Length - d.Width) / 2.0)
    End Function

    Private Function CrossReserveY(ByVal d As TableDesign) As Double
        If d.Constants.MountSketchFixed Then Return 0
        Return Math.Max(0, (d.Width - d.Length) / 2.0)
    End Function

    ' A rail carries one corner bracket at each end, both anchored where the two rails'
    ' side faces meet - at +-(post_center - OffsetWidth) - and each leg runs
    ' GussetSizeCorner inward from there. The two must not run into each other:
    '
    '     post_center - OffsetWidth >= GussetSizeCorner
    '
    ' rearranged into a minimum table size. Touching is allowed, overlapping is not.
    '
    ' The two axes do NOT give the same expression, because post_center_x subtracts
    ' OffsetLength while post_center_y subtracts OffsetWidth. On 45x90 this is the binding
    ' limit in both directions (315 / 270 against 311 / 221 from the plate).
    Private Function CornerBracketMinLength(ByVal d As TableDesign) As Double
        Return 2.0 * d.GussetSizeCorner + 2.0 * d.OffsetLength + 2.0 * d.OffsetWidth
    End Function

    Private Function CornerBracketMinWidth(ByVal d As TableDesign) As Double
        Return 2.0 * d.GussetSizeCorner + 4.0 * d.OffsetWidth
    End Function

    Public Function LengthRange(ByVal d As TableDesign) As ValueRange
        ' The lower plate sits between the length rails and loses one bracket per side
        ' in V2, so the span has to survive that subtraction.
        Dim structural As Double = 4.0 * d.OffsetLength + 2.0 * d.GussetSize +
                                   2.0 * d.Constants.PlateClearance + MinClearPlate
        ' The two corner brackets on one length rail must not overlap.
        Dim corner As Double = CornerBracketMinLength(d)
        ' The robot bolt raster plus the base body has to fit on the plate.
        Dim raster As Double = d.MountPitchX + 2.0 * KeepFromEdge(d)
        Dim min_ As Double = Math.Ceiling(Math.Max(Math.Max(structural, corner), raster))

        Dim why As String
        If corner >= structural AndAlso corner >= raster Then
            why = "Set by the two " & Fmt(d.GussetSizeCorner) &
                  " mm corner brackets meeting on the same rail."
        ElseIf structural >= raster Then
            why = "Set by the " & d.Profile.Caption & " posts and their brackets."
        Else
            why = "Set by the robot base: " & Fmt(d.MountPitchX) & " mm bolt pitch + 2 x " &
                  Fmt(KeepFromEdge(d)) & " mm to the edge."
        End If
        Return New ValueRange(min_, MaxTableSize, why)
    End Function

    Public Function WidthRange(ByVal d As TableDesign) As ValueRange
        Dim structural As Double = 4.0 * d.OffsetWidth + 2.0 * d.GussetSize +
                                   2.0 * d.Constants.PlateClearance + MinClearPlate
        ' Same corner bracket pair, this time on a width rail.
        Dim corner As Double = CornerBracketMinWidth(d)
        Dim raster As Double = d.MountPitchY + 2.0 * KeepFromEdge(d)
        Dim min_ As Double = Math.Ceiling(Math.Max(Math.Max(structural, corner), raster))

        Dim why As String
        If corner >= structural AndAlso corner >= raster Then
            why = "Set by the two " & Fmt(d.GussetSizeCorner) &
                  " mm corner brackets meeting on the same rail."
        ElseIf structural >= raster Then
            why = "Set by the " & d.Profile.Caption & " posts and their brackets."
        Else
            why = "Set by the robot base: " & Fmt(d.MountPitchY) & " mm bolt pitch + 2 x " &
                  Fmt(KeepFromEdge(d)) & " mm to the edge."
        End If
        Return New ValueRange(min_, MaxTableSize, why)
    End Function

    ' Height has to leave room for: feet, the lower rail level (which V1 pushes up by one
    ' bracket), the brackets between the two rail levels, and the worktop plate.
    Public Function HeightRange(ByVal d As TableDesign) As ValueRange
        Dim min_ As Double = Math.Ceiling(d.UpperPlateThickness + d.OffsetLength +
                                          d.HeightLowerHorizontal + MinRailGap(d))
        Dim why As String = "Feet " & Fmt(d.Constants.FeetHeight) & " + lower rail level " &
                            Fmt(d.HeightLowerHorizontal) & " + bracket room + worktop " &
                            Fmt(d.UpperPlateThickness) & " mm."
        Return New ValueRange(min_, MaxTableHeight, why)
    End Function

    ' Two different things cap the worktop, and which one bites depends on the table: the
    ' room the brackets need under it, and the longest M10 the robot bolt can be. The bolt
    ' crosses the base flange, the plate, a washer and a nut and still has to show thread,
    ' so a 70 mm screw runs out at a 42 mm plate - well before the 60 mm the geometry would
    ' still allow. Which of the two bit is the whole point of the reason string: a slider
    ' that merely stops is a slider the user argues with.
    Public Function UpperPlateRange(ByVal d As TableDesign) As ValueRange
        Dim room As Double = Math.Floor(d.Height - d.OffsetLength - d.HeightLowerHorizontal - MinRailGap(d))
        If room > 60 Then room = 60

        Dim bolt As Double = Math.Floor(FastenerCatalog.MaxThickness(FastenerCatalog.RobotMount))
        If bolt < room Then
            Return New ValueRange(5, bolt,
                "Thicker than " & Fmt(bolt) & " mm and no stock M10 reaches through: the longest is " &
                Fmt(FastenerCatalog.M10CapHead.LongestLength) & " mm, of which " &
                Fmt(FastenerCatalog.RobotMountAdd) &
                " mm goes on the base flange, the washer, the nut and the thread past it.")
        End If

        Return New ValueRange(5, room,
                              "Thicker than " & Fmt(room) & " mm would eat the room the brackets need under the worktop.")
    End Function

    Public Function LowerPlateRange(ByVal d As TableDesign) As ValueRange
        Dim max_ As Double = Math.Floor(d.RailGap - 2.0 * d.OffsetLength - 10.0)
        If max_ > 40 Then max_ = 40
        If max_ < 2 Then max_ = 2
        Return New ValueRange(2, max_, "The shelf must stay clear of the upper rails.")
    End Function

    ' --- robot base offsets ---
    ' This is the input that used to break the update: too small an offset pushed the
    ' outer bolt hole through the edge of the plate.

    ' The offset is measured from the plate edge at +X to the first bolt, and the raster
    ' grows from there towards -X. So the range runs the whole plate: at the minimum the
    ' base sits against the +X edge, at the maximum against the -X edge. Anything in
    ' between is reachable, which is what stops the robot from being stuck in one quadrant.
    Public Function OffsetXRange(ByVal d As TableDesign) As ValueRange
        Dim keep As Double = KeepFromEdge(d)
        Dim min_ As Double = Math.Ceiling(keep + CrossReserveX(d))
        Dim max_ As Double = Math.Floor(d.Length - keep - d.MountPitchX)
        Return New ValueRange(min_, max_, OffsetWhy(d, min_, max_, CrossReserveX(d), "across the width"))
    End Function

    Public Function OffsetYRange(ByVal d As TableDesign) As ValueRange
        Dim keep As Double = KeepFromEdge(d)
        Dim min_ As Double = Math.Ceiling(keep + CrossReserveY(d))
        Dim max_ As Double = Math.Floor(d.Width - keep - d.MountPitchY)
        Return New ValueRange(min_, max_, OffsetWhy(d, min_, max_, CrossReserveY(d), "along the length"))
    End Function

    Private Function OffsetWhy(ByVal d As TableDesign, ByVal min_ As Double, ByVal max_ As Double,
                               ByVal reserve As Double, ByVal otherAxis As String) As String
        If reserve > 0 Then
            Return "Below " & Fmt(min_) & " mm the base leaves the plate - the master still applies this " &
                   otherAxis & "."
        End If
        Return "From " & Fmt(min_) & " mm (base against this edge) to " & Fmt(max_) &
               " mm (against the far edge)."
    End Function

    ' --- robot arm (task: nominal .. nominal + 20 %) ---

    Public Function ArmJ2J3Range(ByVal d As TableDesign) As ValueRange
        Dim n As Double = d.Constants.NominalLenJ2J3
        Return New ValueRange(n, Math.Floor(n * 1.2),
                              "The task allows the nominal RoboDK length up to +20 %.")
    End Function

    Public Function ArmJ3J5Range(ByVal d As TableDesign) As ValueRange
        Dim n As Double = d.Constants.NominalLenJ3J5
        Return New ValueRange(n, Math.Floor(n * 1.2),
                              "The task allows the nominal RoboDK length up to +20 %.")
    End Function

    ' --- path points (optional, section 12) ---

    ' EVERY path coordinate is in the ROBOT BASE FRAME - origin on the robot's mounting
    ' face, X forward, Z up.
    '
    ' They used to be phrased in TABLE terms (+-Length/2, +-Width/2, floor-referenced
    ' height) while being consumed as robot-frame values. That let P1X go to zero or below,
    ' which RobotKinematics.Solve refuses every single time ("not in front of the robot"),
    ' and started P1Z at the worktop height for a number measured from the mounting face.
    ' Fixed 29.08.2026 together with the P3/P6 inputs.

    ' Along the robot's X axis, both ways. Behind the robot is allowed since 24.09.2026,
    ' when J1 moved to the half-angle form (RobotKinematics.SolveAt) - atan(y/x) needed
    ' x > 0. The bound keeps a tenth of the reach in hand so the elbow never has to lock
    ' straight. A point too close to the J1 axis (inside the fold) is not something a
    ' range on X alone can express; SolveAt refuses it and the checks panel says so.
    Public Function PathForwardRange(ByVal d As TableDesign) As ValueRange
        Dim p As RobotKinematics.PathInputs = d.ToPathInputs()
        Dim reach As Double = RobotKinematics.Reach(p)
        Dim lim As Double = Math.Floor(reach * 0.9)
        Return New ValueRange(-lim, lim,
            "Along the robot's own X axis, negative behind it. It reaches " & Fmt(reach) &
            " mm; the last tenth is kept in hand so the elbow never locks straight.")
    End Function

    ' Sideways. Bounded by the same reach, since a point far to the side is just as far
    ' from the shoulder as one straight ahead.
    Public Function PathSidewaysRange(ByVal d As TableDesign) As ValueRange
        Dim p As RobotKinematics.PathInputs = d.ToPathInputs()
        Dim half As Double = Math.Floor(RobotKinematics.Reach(p) * 0.6)
        Return New ValueRange(-half, half,
            "Sideways from the robot's X axis, in its own frame. Positive is the side the " &
            "place station moves toward.")
    End Function

    ' Height above the robot's MOUNTING FACE, not the floor. Named apart from HeightRange,
    ' which is the table's own height and measured from the floor.
    Public Function PathHeightRange(ByVal d As TableDesign) As ValueRange
        Dim p As RobotKinematics.PathInputs = d.ToPathInputs()
        Return New ValueRange(0, Math.Floor(RobotKinematics.RB_D1 + RobotKinematics.Reach(p) * 0.8),
            "Above the robot's mounting face, so 0 is the plate it stands on.")
    End Function

    Public Function SegmentRange(ByVal maxLen As Double) As ValueRange
        Return New ValueRange(10, maxLen, "Path segment length.")
    End Function

    ' --- how fast the simulation runs -----------------------------------------------------

    ' Seconds per MoveJ leg - the two free-space moves, P1->P2 and P4->P6.
    Public Function SimLegRange() As ValueRange
        Return New ValueRange(0.2, 10,
            "Seconds for each of the two joint moves through free space.")
    End Function

    ' Tool speed on the MoveL legs, everything between P2 and P5. A real LR Mate does up to
    ' about 2 m/s; the ceiling here is well under that because a simulation people have to
    ' watch is not improved by being fast.
    Public Function SimSpeedRange() As ValueRange
        Return New ValueRange(10, 500,
            "Tool-tip speed on the straight legs, mm/s. The joint legs use their own time.")
    End Function

    ' --- the tool ---------------------------------------------------------------------------

    ' Thin enough and the M5 flange screws have nothing to hold; thick enough and the
    ' tool is a lever. Neither bound is a catalogue figure, so both are stated as
    ' engineering judgement rather than dressed up as one.
    Public Function AdapterRange() As ValueRange
        Return New ValueRange(6, 40,
            "Below 6 mm the flange screws have no thread engagement; above 40 the plate " &
            "is mostly a lever arm.")
    End Function

    Public Function BaseOverhangRange() As ValueRange
        Return New ValueRange(0, 150, "Read it off the robot base; 0 checks the bolt holes only.")
    End Function

    ' --- robot bolt raster ---

    ' A pitch has to leave room for the hole itself plus the edge distance on both sides.
    ' Both pitches are bounded by the SMALLER table dimension, because the mount variant
    ' may put either one on either axis and switching it must never break the table.
    Public Function MountPitchRange(ByVal d As TableDesign) As ValueRange
        Dim min_ As Double = Math.Ceiling(Math.Max(2.0 * d.MountHoleDiameter, 20.0))
        Dim max_ As Double = Math.Floor(Math.Min(d.Length, d.Width) - 2.0 * KeepFromEdge(d))
        If max_ < min_ Then max_ = min_
        Return New ValueRange(min_, max_,
            "Bolt raster of the robot base. Bounded by the " & Fmt(Math.Min(d.Length, d.Width)) &
            " mm side less 2 x " & Fmt(KeepFromEdge(d)) & " mm to the edge.")
    End Function

    ' The hole is a clearance hole, so it is read off the bolt, not chosen freely. The cap
    ' keeps it from swallowing the raster it sits in.
    Public Function MountHoleRange(ByVal d As TableDesign) As ValueRange
        Dim smallest As Double = Math.Min(d.MountPitchLong, d.MountPitchShort)
        Dim max_ As Double = Math.Floor(Math.Min(30.0, smallest / 2.0))
        If max_ < 3 Then max_ = 3
        Return New ValueRange(3, max_, "Clearance hole for the base bolt: 9 for M8, 11 for M10, 14 for M12.")
    End Function

    ' Room the two bracket rows need between the rail levels.
    Private Function MinRailGap(ByVal d As TableDesign) As Double
        Return 2.0 * d.GussetSize + MinRailHeadroom
    End Function

    ' --- cross checks the sliders cannot express on their own ---

    Public Function Review(ByVal d As TableDesign) As List(Of DesignCheck)
        Dim list As New List(Of DesignCheck)()

        ' Bolt hole landing on a corner post: legal, but the post head and the robot foot
        ' then fight over the same spot on the worktop.
        Dim clash As Boolean = False
        For Each hx As Double In d.MountHolesX()
            For Each hy As Double In d.MountHolesY()
                If Math.Abs(Math.Abs(hx) - d.PostCenterX) <= d.OffsetLength AndAlso
                   Math.Abs(Math.Abs(hy) - d.PostCenterY) <= d.OffsetWidth Then
                    clash = True
                End If
            Next
        Next
        If clash Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                "A robot base bolt sits right over a corner post. Move the robot further from the edge."))
        End If

        ' How much plate is left around the raster. Both edges have to be checked, not
        ' just the one the offset is measured from - the base may sit past the middle.
        Dim keep As Double = KeepFromEdge(d)
        Dim marginX As Double = EdgeMargin(d.MountHolesX(), d.Length / 2.0)
        Dim marginY As Double = EdgeMargin(d.MountHolesY(), d.Width / 2.0)
        Dim tightest As Double = Math.Min(marginX, marginY)
        Dim clearance As Double = tightest - d.BaseOverhang   ' plate edge -> base body edge
        If tightest < d.Constants.HoleEdgeDistance Then
            list.Add(New DesignCheck(CheckLevel.Problem,
                "Only " & Fmt(tightest) & " mm of plate left around the outer robot bolt - the update will fail."))
        ElseIf clearance < 0 Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                "The robot base hangs " & Fmt(-clearance) & " mm over the plate edge."))
        ElseIf clearance < 10 Then
            list.Add(New DesignCheck(CheckLevel.Info,
                "Robot base edge is " & Fmt(clearance) & " mm inside the plate - as tight as the limits allow."))
        End If

        ' Where the raster came from. An unverified figure is the one thing here the model
        ' cannot catch on its own: the holes get drilled wherever the number says.
        Dim rb As RobotInfo = d.Robot
        If Not rb.HasRaster Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                rb.Caption & " has no recorded bolt raster - the numbers below are yours, not " &
                "the robot's. Confirm them against the manufacturer drawing (see robots.txt)."))
        ElseIf Not rb.Verified Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                rb.Caption & ": the raster is UNVERIFIED - nobody has traced it to a " &
                "manufacturer drawing yet (see robots.txt)."))
        ElseIf Not d.MatchesCatalogue Then
            list.Add(New DesignCheck(CheckLevel.Info,
                "Raster edited by hand: " & Fmt(d.MountPitchLong) & " x " & Fmt(d.MountPitchShort) &
                ", hole " & Fmt(d.MountHoleDiameter) & " mm, against the robot's " &
                Fmt(rb.BoltPitchLong) & " x " & Fmt(rb.BoltPitchShort) & ", hole " &
                Fmt(rb.HoleDiameter) & " mm."))
        End If

        ' The master's hole sketch is rotated 90 degrees against the part axes, so the two
        ' offsets swap axes on their way into the geometry. Saying it out loud beats
        ' silently compensating: the reserve in the minimums is only a stopgap.
        If Not d.Constants.MountSketchFixed Then
            Dim reserve As Double = Math.Max(CrossReserveX(d), CrossReserveY(d))
            Dim what As String
            If reserve > 0 Then
                what = "the minimums reserve " & Fmt(reserve) & " mm extra"
            Else
                what = "the plate is square enough that no extra reserve is needed"
            End If
            list.Add(New DesignCheck(CheckLevel.Warning,
                "Master applies the base offsets rotated 90 deg, so " & what &
                ". Fix the hole sketch, then set RobotMountSketchFixed=true."))
        End If

        ' The shelf: V2 loses a bracket per side.
        If d.LowerPlateLength < MinClearPlate Then
            list.Add(New DesignCheck(CheckLevel.Problem,
                "The shelf would end up only " & Fmt(d.LowerPlateLength) & " mm long. Make the table longer."))
        End If

        If d.RailGap < MinRailGap(d) Then
            list.Add(New DesignCheck(CheckLevel.Problem,
                "Only " & Fmt(d.RailGap) & " mm between the rail levels; the " & Fmt(d.GussetSize) &
                " mm brackets do not fit. Raise the table."))
        End If

        ' Say out loud that this one needs Apply - the master alone will not do it.
        If d.Simplified Then
            list.Add(New DesignCheck(CheckLevel.Info,
                "Simplified swaps the join onto the 'profile_simple' surface, which only " &
                "happens on Apply; changing the parameter by hand in CATIA does not do it."))
        End If

        If d.Length / d.Width > 3.0 OrElse d.Width / d.Length > 3.0 Then
            list.Add(New DesignCheck(CheckLevel.Info,
                "The table is unusually long and narrow - check that it still stands stably."))
        End If

        list.AddRange(FastenerReview(d))

        Return list
    End Function

    ' --- screws ---
    '
    ' These never break a CATIA update, so none of them is a Problem: a 3 mm shelf with a
    ' countersunk M8 rebuilds perfectly happily and is simply not buildable in metal. The
    ' panel is where that gets said out loud, instead of the drawing quietly claiming it.
    Private Function FastenerReview(ByVal d As TableDesign) As List(Of DesignCheck)
        Dim list As New List(Of DesignCheck)()
        Dim joints As List(Of FastenerChoice) = FastenerCatalog.Plan(d)

        For Each c As FastenerChoice In joints
            ' The robot bolt fails in its own ways: no countersink to hide a head in, and
            ' nothing under the nut to bottom out on. Running it through the checks written
            ' for the countersunk plate screws would produce three sentences that are all
            ' about the wrong screw.
            If c.Joint Is FastenerCatalog.RobotMount Then
                RobotMountReview(c, list)
            ElseIf c.Thickness < FastenerCatalog.MinPlateThickness Then
                list.Add(New DesignCheck(CheckLevel.Warning,
                    c.Joint.Caption & ": a " & Fmt(c.Thickness) &
                    " mm plate cannot hide a DIN 7991 head (" & Fmt(FastenerCatalog.HeadHeight) &
                    " mm deep). Use at least " & Fmt(FastenerCatalog.MinPlateThickness) & " mm."))
            ElseIf c.Length <= 0 Then
                list.Add(New DesignCheck(CheckLevel.Warning,
                    c.Joint.Caption & ": a " & Fmt(c.Thickness) & " mm plate would need an M8 x " &
                    Fmt(c.Thickness + c.Joint.MinReach) & ", longer than any stock DIN 7991. " &
                    "Thickest plate this joint takes is " &
                    Fmt(FastenerCatalog.MaxThickness(c.Joint)) & " mm."))
            ElseIf c.TooDeep Then
                list.Add(New DesignCheck(CheckLevel.Warning,
                    c.Joint.Caption & ": the next stock size is " & c.Designation & ", whose tip goes " &
                    Fmt(c.Reach) & " mm past the plate - " & Fmt(c.Joint.MaxReach) &
                    " mm is all the room there is. Change the plate by a millimetre or two and " &
                    "the shorter screw fits."))
            End If
        Next

        ' What to order. The lengths follow the plate thickness on their own, so the only
        ' place they are ever visible before the BOM is here. Three plate joints, three
        ' lengths: the worktop's two cannot share one, because the tapped core hole takes a
        ' screw that would bottom out in the slot next to it.
        Dim shopping As New System.Text.StringBuilder()
        For Each c As FastenerChoice In joints
            If c.Length > 0 AndAlso Not c.Joint Is FastenerCatalog.RobotMount Then
                If shopping.Length > 0 Then shopping.Append(", ")
                shopping.Append(c.Joint.Key & " " & c.Designation)
            End If
        Next
        If shopping.Length > 0 Then
            list.Add(New DesignCheck(CheckLevel.Info, "Plate screws: " & shopping.ToString() & "."))
        End If

        ' The robot base gets its own line, because it is the only joint that orders three
        ' different articles at once - and because two of the three never change, so a
        ' reader who sees them move knows something is wrong.
        Dim rm As FastenerChoice = RobotMountChoice(joints)
        If rm IsNot Nothing AndAlso rm.Length > 0 Then
            list.Add(New DesignCheck(CheckLevel.Info,
                "Robot base: " & FastenerCatalog.RobotMountCount & " x " & rm.Designation &
                ", " & FastenerCatalog.RobotMountCount & " x " & FastenerCatalog.WasherM10 &
                ", " & FastenerCatalog.RobotMountCount & " x " & FastenerCatalog.NutM10 & "."))
        End If

        ReviewPath(d, list)
        ReviewTool(d, list)

        Return list
    End Function

    Private Function RobotMountChoice(ByVal joints As List(Of FastenerChoice)) As FastenerChoice
        For Each c As FastenerChoice In joints
            If c.Joint Is FastenerCatalog.RobotMount Then Return c
        Next
        Return Nothing
    End Function

    ' The robot base bolt: DIN 912 through the base flange, the worktop, a washer and a nut.
    '
    ' Only two things can go wrong, and neither is the same failure the plate screws have.
    ' The series can run out - the plate got thicker than a 70 mm bolt can cross - which the
    ' worktop slider already prevents, so reaching this branch means the master was edited
    ' behind the app's back. Or the plate is thin enough that even the shortest stock bolt
    ' leaves a stub of thread hanging under the table: harmless, but it belongs in a drawing
    ' nobody has to explain, so it is said here rather than discovered on the bench.
    Private Sub RobotMountReview(ByVal c As FastenerChoice, ByVal list As List(Of DesignCheck))
        If c.Length <= 0 Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                c.Joint.Caption & ": a " & Fmt(c.Thickness) & " mm worktop would need an M10 x " &
                Fmt(c.Thickness + c.Joint.MinReach) & ", longer than any stock DIN 912. " &
                "Thickest worktop this joint takes is " &
                Fmt(FastenerCatalog.MaxThickness(c.Joint)) & " mm."))
        ElseIf c.TooDeep Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                c.Joint.Caption & ": the shortest stock bolt is " & c.Designation &
                ", which leaves " & Fmt(c.Reach - FastenerCatalog.RobotMountGrip) &
                " mm of thread below the nut. Nothing is in its way, but a thicker worktop " &
                "or a shorter bolt would look like it was meant."))
        End If
    End Sub


    ' --- the path -----------------------------------------------------------------------
    '
    ' The sliders already refuse an out-of-range number, but they each see one value at a
    ' time. Whether the ARM can actually get to a point is a question about all of them at
    ' once, and it is worth answering here rather than in CATIA: an unreachable pose shows
    ' up there as a broken relation or a mechanism that refuses to move, neither of which
    ' says "P3 is 40 mm too far away".
    Private Sub ReviewPath(ByVal d As TableDesign, ByVal list As List(Of DesignCheck))
        Dim p As RobotKinematics.PathInputs = d.ToPathInputs()

        Dim reach As Double = RobotKinematics.Reach(p)
        list.Add(New DesignCheck(CheckLevel.Info,
            "Arm reach " & Fmt(reach) & " mm, tool " & Fmt(p.ToolLen) &
            " mm from the flange to the grip (" & d.ToolLenSource() & ")."))

        Dim worst As Double = 0
        Dim bad As Integer = 0
        For Each name As String In New String() {"P1", "P2", "P3", "P4", "P5", "P6"}
            Dim pose As RobotPose = RobotKinematics.Solve(p, name)
            If Not pose.Ok Then
                bad += 1
                list.Add(New DesignCheck(CheckLevel.Problem, name & ": " & pose.Problem))
            Else
                ' Round trip: put the angles back through the forward chain. If this
                ' ever disagrees, the two copies of the kinematics - here and in the
                ' master's formulas - have drifted apart, and everything downstream is
                ' quietly wrong.
                Dim got As Double() = RobotKinematics.Forward(pose, p)
                Dim want As Double() = RobotKinematics.Point(p, name)
                Dim e As Double = Math.Sqrt((got(0) - want(0)) ^ 2 +
                                            (got(1) - want(1)) ^ 2 +
                                            (got(2) - want(2)) ^ 2)
                If e > worst Then worst = e
            End If
        Next

        If bad > 0 Then Exit Sub

        ' What this can and cannot accuse. Solve and Forward are BOTH in this program, so a
        ' disagreement is a bug HERE - it says nothing about the master's formulas or the
        ' mechanism. The first version read "the angle convention is wrong somewhere", which
        ' sent the reader to CATIA to look for a fault that was in this file: Forward was
        ' still subtracting tool_len alone after Solve had been corrected to tool_len + rb_d6,
        ' so it reported exactly 80 mm on a model that was right (30.08.2026).
        If worst > 0.01 Then
            list.Add(New DesignCheck(CheckLevel.Problem,
                "This window's kinematics disagree with themselves by " & Fmt(worst) &
                " mm (Solve vs Forward, both in RobotKinematics.vb). The model in CATIA is " &
                "not what this is about - check tool_len + rb_d6 here first."))
            Exit Sub
        End If

        ' The six key points being reachable is NOT enough once the legs between P2 and P5
        ' are travelled in a straight line. A line between two reachable points can still
        ' leave the workspace or pass over the shoulder halfway along, and under MoveL the
        ' arm really does go there - so the samples the driver will play are tested too.
        Dim samples As List(Of RobotPose) = RobotKinematics.SampledPath(p, PATH_TOL_MM)
        Dim firstBad As RobotPose = Nothing
        Dim nBad As Integer = 0
        For Each s As RobotPose In samples
            If Not s.Ok Then
                nBad += 1
                If firstBad Is Nothing Then firstBad = s
            End If
        Next

        If nBad > 0 Then
            list.Add(New DesignCheck(CheckLevel.Problem,
                "The six points are reachable, but the arm cannot follow the straight " &
                "legs between them: " & firstBad.Name & " - " & firstBad.Problem &
                If(nBad > 1, "  (and " & (nBad - 1) & " more samples)", "")))
            Exit Sub
        End If

        list.Add(New DesignCheck(CheckLevel.Info,
            "All six points reachable and every straight leg clear, " & samples.Count &
            " poses checked, joints inside travel."))
    End Sub

    ' How straight a MoveL leg has to be, in millimetres. A real robot holds a linear path
    ' to about this; the catalogue +-0.02 mm for the LR Mate is repeatability of returning
    ' to a taught point, not straightness between them.
    Public Const PATH_TOL_MM As Double = 0.1


    ' --- the tool -----------------------------------------------------------------------
    Private Sub ReviewTool(ByVal d As TableDesign, ByVal list As List(Of DesignCheck))
        Dim g As GripperInfo = d.Gripper

        Dim why As String = GripperCatalog.GripProblem(g, d.FingerReach())
        If why <> "" Then
            list.Add(New DesignCheck(CheckLevel.Problem, "Gripper: " & why & "."))
        Else
            list.Add(New DesignCheck(CheckLevel.Info,
                g.Caption & " on a " & Fmt(g.ObjA) & " x " & Fmt(g.ObjB) & " x " &
                Fmt(g.ObjC) & " mm object; " & Fmt(g.TotalStroke) &
                " mm total jaw travel."))
        End If

        ' Same contract as RobotCatalog: an unverified figure is said out loud rather
        ' than shown as if it were measured.
        If Not g.Verified Then
            list.Add(New DesignCheck(CheckLevel.Warning,
                "The " & g.Caption & " figures come from the catalogue, not from the " &
                "CAD. Measure the mounting and finger interfaces before drilling an " &
                "adapter plate - TODO2.md M3 and M4."))
        End If
    End Sub

    ' Smallest distance from either edge (+half / -half) to the outermost hole.
    Private Function EdgeMargin(ByVal holes As Double(), ByVal half As Double) As Double
        Dim lo As Double = holes(0), hi As Double = holes(0)
        For Each h As Double In holes
            If h < lo Then lo = h
            If h > hi Then hi = h
        Next
        Return Math.Min(half - hi, lo + half)
    End Function

    Public Function Fmt(ByVal v As Double) As String
        Return v.ToString("0.##", Globalization.CultureInfo.InvariantCulture)
    End Function

End Module
