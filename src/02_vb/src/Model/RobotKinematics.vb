Option Explicit On
Option Strict Off

Imports System.Math

' Inverse kinematics of the FANUC LR Mate 200iC, mirrored from the model.
'
' This is to the master's IK formulas what ProfileCatalog is to the rule R_Profile and
' FastenerCatalog is to R_Fasteners: a deliberate second copy, in the application, so
' the window can show the user what the arm will do BEFORE anything is written to CATIA
' and before an update can fail. The reference implementation of all three copies -
' this one, the Knowledge formulas, and the check the validator runs - is
' tools/ik_oracle.py, where the algebra is verified by putting every solution back
' through the forward chain (worst round-trip error 0.000000 mm).
'
' It had a second job until 30.08.2026 - channel B of TODO2.md phase 15, where
' MechanismDriver pushed the poses computed here straight into the mechanism. That channel
' is gone with the Simulate button: the Knowledge laws drive the model, and two drivers on
' one mechanism cannot both hold it. What is left is the job above, which is the one that
' matters here - the window says whether the arm can reach a point BEFORE anything is
' written to CATIA.
'
' WHY THE PROBLEM IS SOLVABLE IN CLOSED FORM
' The task states that the tool's rotation angles need not be controlled. So the tool
' is kept pointing straight down, which fixes J4 = J6 = 0, makes J5 follow from J2 and
' J3, and - the part that matters - puts the wrist centre exactly ToolLen above the
' tool tip. The arm therefore has to reach a POINT, not a frame: one rotation about the
' base, then a planar two-link problem solved by the cosine rule.
'
' THE FOREARM IS NOT STRAIGHT. J3 to J5 is the hypotenuse of a3 and d4, tilted by
' phi3 = atan(a3/d4) away from the d4 direction. Forgetting that tilt produces an arm
' that is exactly right in the zero pose and wrong in every other one - which is the
' kind of error that survives a casual look at the screen.
'
' ANGLE CONVENTION, stated once and obeyed everywhere: all angles are DMU command
' values, and all are zero in the pose the vendor CAD is modelled in - arm straight up,
' forearm straight out along +X. That makes "all commands zero" the home position the
' task asks to match, with no offset table anywhere.

Public Class RobotPose

    Public ReadOnly Name As String
    Public ReadOnly J1 As Double
    Public ReadOnly J2 As Double
    Public ReadOnly J3 As Double
    Public ReadOnly J4 As Double
    Public ReadOnly J5 As Double
    Public ReadOnly J6 As Double

    ' Empty when the pose is fine; otherwise says which joint is over its travel, or
    ' why the point cannot be reached at all.
    Public ReadOnly Problem As String

    Public Sub New(ByVal name_ As String,
                   ByVal j1_ As Double, ByVal j2_ As Double, ByVal j3_ As Double,
                   ByVal j4_ As Double, ByVal j5_ As Double, ByVal j6_ As Double,
                   ByVal problem_ As String)
        Name = name_
        J1 = j1_ : J2 = j2_ : J3 = j3_ : J4 = j4_ : J5 = j5_ : J6 = j6_
        Problem = problem_
    End Sub

    Public ReadOnly Property Ok As Boolean
        Get
            Return Problem = ""
        End Get
    End Property

    ' The four values a DMU command set needs, in command order: J1, J2, J3, J5.
    Public Function Driven() As Double()
        Return New Double() {J1, J2, J3, J5}
    End Function

End Class


Public Module RobotKinematics

    ' --- machine constants, from the DH table -------------------------------------------
    ' Recovered twice and independently: by decoding
    ' src/11_roboDK/Fanuc-LR-Mate-200iC.robot, and from the RoboDK screenshot embedded
    ' in TASK/Zadatak_ROMS_2024_2025_Kinematika.xlsx. They agree exactly.
    Public Const RB_D1 As Double = 330.0   ' mounting face to the J2 axis
    Public Const RB_A1 As Double = 75.0    ' J1 axis to J2 axis, horizontal
    Public Const RB_A3 As Double = 75.0    ' J3 to J5, perpendicular offset
    Public Const RB_D6 As Double = 80.0    ' J5 to the tool flange face

    Public Const NOM_J2J3 As Double = 300.0
    Public Const NOM_J3J5 As Double = 320.0

    ' The task allows nominal up to +20 %.
    Public Const MAX_STRETCH As Double = 1.2

    ' Joint travel, degrees, RoboDK's decoupled convention. Same numbers go into the
    ' DMU joint limits, where "Check joint limits" enforces them during the replay.
    Public ReadOnly JointMin As Double() = {-170.0, -60.0, -70.0, -190.0, -115.0, -400.0}
    Public ReadOnly JointMax As Double() = {170.0, 140.0, 170.0, 190.0, 115.0, 400.0}

    ' The eight key frames, in the order the tool visits them. P2 and P4 appear twice -
    ' that is what makes it a pick-and-place rather than a tour of six points.
    Public ReadOnly Frames As String() =
        {"P1", "P2", "P3", "P2", "P4", "P5", "P4", "P6"}

    ' HOW each leg is travelled, one entry per leg, so Legs(k) covers Frames(k)->Frames(k+1).
    '
    ' True means MoveL: the TIP goes in a straight line, which is what the task asks for on
    ' everything between P2 and P5. False means MoveJ: the ANGLES interpolate and the tip
    ' swings, which is how a robot crosses free space.
    '
    ' The distinction is not cosmetic. Measured on this model, P2->P4 under MoveJ deviates
    ' 7.7 mm from the straight line: both ends share J2/J3/J5 (same radius, same height), so
    ' only J1 moves and the tip traces an exact circular arc of radius 419.7. The descents
    ' deviate 1.55 mm.
    Public ReadOnly LegIsLinear As Boolean() =
        {False, True, True, True, True, True, False}


    ' Everything the window can change that affects the path.
    Public Class PathInputs
        Public LenJ2J3 As Double = NOM_J2J3
        Public LenJ3J5 As Double = NOM_J3J5
        ' P1, P3 and P6 are typed directly; P2, P4 and P5 follow from P3 and the two
        ' segment lengths. See Point().
        Public P1X As Double = 275.0
        Public P1Y As Double = 60.0
        Public P1Z As Double = 600.0
        Public P3X As Double = 412.0
        Public P3Y As Double = -80.0
        Public P3Z As Double = 150.0
        Public P6X As Double = 275.0
        Public P6Y As Double = -60.0
        Public P6Z As Double = 600.0
        Public P2P3 As Double = 60.0
        Public P2P4 As Double = 160.0
        ' Flange face to the gripping point: adapter plate + gripper body + finger.
        Public ToolLen As Double = 96.0
        ' Half the object's height - where the tool grips it.
        Public GripZ As Double = 30.0

        ' P4P5 is not an input. The task figure puts both stations at one height, so the
        ' descent at the place station equals the descent at the pick station; an
        ' independent value would mean one station floats higher than the other.
        Public ReadOnly Property P4P5 As Double
            Get
                Return P2P3
            End Get
        End Property
    End Class


    ' --- derived -------------------------------------------------------------------------

    ' J3 to J5, straight line. The forearm is offset, so this is not LenJ3J5.
    Public Function L3(ByVal p As PathInputs) As Double
        Return Sqrt(RB_A3 * RB_A3 + p.LenJ3J5 * p.LenJ3J5)
    End Function

    ' Tilt of the J3-J5 line away from the d4 direction, degrees.
    Public Function Phi3(ByVal p As PathInputs) As Double
        Return Deg(Atan(RB_A3 / p.LenJ3J5))
    End Function

    ' Elbow angle in the CAD pose, so the J3 command reads zero there.
    Public Function J3Zero(ByVal p As PathInputs) As Double
        Return 90.0 - Phi3(p)
    End Function

    Public Function Reach(ByVal p As PathInputs) As Double
        Return RB_A1 + p.LenJ2J3 + L3(p)
    End Function

    ' --- the six points, in the robot base frame -----------------------------------------

    ' The path is a RECTANGLE, not an arc. Read straight off the task figure: P2P4 is a
    ' plain offset along Y (the robot's short axis) and P2P3 one along Z, with all four
    ' station points sharing x = P3X.
    '
    '     P2 ------ P2P4 (Y) ------ P4          P3 = input, the cube sits here
    '      |                         |
    '   P2P3 (Z)                  = P2P3
    '      |                         |
    '     P3                        P5
    '
    ' An earlier version placed both stations on one circle about the J1 axis, which made
    ' P2P4 a chord and therefore NOT independent of P3Y. That circle was a modelling
    ' decision, not something the task asks for, and dropping it removes the whole polar
    ' layer - path_radius, path_phi, station_height and path_pick_z all went with it.
    Public Function Point(ByVal p As PathInputs, ByVal name As String) As Double()
        Dim y4 As Double = p.P3Y + p.P2P4
        Dim zTop As Double = p.P3Z + p.P2P3

        Select Case name
            Case "P1" : Return New Double() {p.P1X, p.P1Y, p.P1Z}
            Case "P2" : Return New Double() {p.P3X, p.P3Y, zTop}
            Case "P3" : Return New Double() {p.P3X, p.P3Y, p.P3Z}
            Case "P4" : Return New Double() {p.P3X, y4, zTop}
            Case "P5" : Return New Double() {p.P3X, y4, p.P3Z}
            Case "P6" : Return New Double() {p.P6X, p.P6Y, p.P6Z}
        End Select
        Return New Double() {0, 0, 0}
    End Function


    ' A point a fraction `s` (0..1) of the way along the straight line between two frames.
    ' This is what makes MoveL straight: the TIP is interpolated, then the inverse
    ' kinematics is solved there, instead of interpolating the angles and letting the tip
    ' go where it likes.
    Public Function Lerp(ByVal a As Double(), ByVal b As Double(),
                         ByVal s As Double) As Double()
        Return New Double() {a(0) + (b(0) - a(0)) * s,
                             a(1) + (b(1) - a(1)) * s,
                             a(2) + (b(2) - a(2)) * s}
    End Function


    ' How many sub-steps a straight leg needs so the tip never strays further than
    ' `tolMm` from the line.
    '
    ' Between two solved samples the mechanism still interpolates angles, so each sub-leg
    ' carries the same arc error as a whole leg would, scaled down: the sagitta of a chord
    ' `c` on radius `R` is about c^2/(8R). Solving for c gives the longest chord allowed,
    ' and the count follows from the leg's length. R is taken as the reach, which is the
    ' smallest curvature the arm can produce - so the answer errs safe.
    Public Function LinearSteps(ByVal p As PathInputs, ByVal a As Double(),
                                ByVal b As Double(), ByVal tolMm As Double) As Integer
        Dim len As Double = Distance(a, b)
        If len <= 0.000001 Then Return 1
        If tolMm <= 0.0 Then tolMm = 0.1

        Dim maxChord As Double = Sqrt(8.0 * Reach(p) * tolMm)
        If maxChord <= 0.000001 Then Return 1

        Dim n As Integer = CInt(Math.Ceiling(len / maxChord))
        If n < 1 Then n = 1
        Return n
    End Function


    Public Function Distance(ByVal a As Double(), ByVal b As Double()) As Double
        Dim dx As Double = b(0) - a(0)
        Dim dy As Double = b(1) - a(1)
        Dim dz As Double = b(2) - a(2)
        Return Sqrt(dx * dx + dy * dy + dz * dz)
    End Function


    ' --- the solution ----------------------------------------------------------------------

    ' Joint angles that put the tool tip on `name`, pointing straight down.
    '
    ' Never throws. An unreachable point comes back as a RobotPose whose Problem says
    ' why, because this runs on every keystroke in the window and an exception there
    ' would be a dialog box per character.
    Public Function Solve(ByVal p As PathInputs, ByVal name As String) As RobotPose
        Return SolveAt(p, name, Point(p, name))
    End Function


    ' The same solution at an ARBITRARY point, which is what MoveL needs: a straight leg is
    ' sampled along the line and this is called once per sample. `name` is only carried
    ' through so a refusal can say where on the path it happened.
    Public Function SolveAt(ByVal p As PathInputs, ByVal name As String,
                            ByVal t As Double()) As RobotPose

        Dim x As Double = t(0), y As Double = t(1), z As Double = t(2)

        ' The Knowledge parser on this CATIA release has no atan2, so the model cannot use
        ' one either, and the two copies must agree. Until 24.09.2026 both used atan(y/x),
        ' which is only right in front of the robot, so x > 0 was enforced. The half-angle
        ' form 2*atan(y / (r + x)) IS atan2 for every direction but one - straight behind
        ' (y = 0, x < 0), where r + x = 0 - and that one is outside J1's +-170 travel anyway.
        Dim r As Double = Sqrt(x * x + y * y)
        If r + x <= 0.000001 Then
            Return Bad(name, "point is straight behind the robot (J1 would be 180 deg)")
        End If

        Dim j1 As Double = Deg(2.0 * Atan(y / (r + x)))

        ' The wrist centre is tool_len + rb_d6 above the tip.
        '
        ' The two-link chain reaches J5, the WRIST CENTRE, because rb_L3 is the J3-to-J5 line.
        ' The flange sits rb_d6 = 80 mm beyond it, and tool_len is measured from the FLANGE
        ' (TOOL_AXS) to the grip. Adding only tool_len put the wrist 80 mm low and the tip went
        ' 80 mm through every target - measured in CATIA on 30.08.2026, exactly 80 mm.
        Dim zw As Double = z + p.ToolLen + RB_D6

        Dim dr As Double = r - RB_A1
        Dim dz As Double = zw - RB_D1
        Dim dist As Double = Sqrt(dr * dr + dz * dz)

        Dim l2 As Double = p.LenJ2J3
        Dim l3v As Double = L3(p)

        If dist > l2 + l3v Then
            Return Bad(name, "too far: " & Fmt(dist) & " mm from the shoulder, the arm spans " &
                             Fmt(l2 + l3v))
        End If
        If dist < Abs(l2 - l3v) Then
            Return Bad(name, "too close: " & Fmt(dist) & " mm from the shoulder, the arm " &
                             "cannot fold below " & Fmt(Abs(l2 - l3v)))
        End If
        If dr <= 0.000001 Then
            Return Bad(name, "the point sits over the shoulder; keep it further out from the J1 axis")
        End If

        Dim psi As Double = Deg(Atan(dz / dr))     ' from horizontal
        Dim toWrist As Double = 90.0 - psi         ' from vertical

        Dim ca As Double = (l2 * l2 + dist * dist - l3v * l3v) / (2.0 * l2 * dist)
        Dim cb As Double = (l2 * l2 + l3v * l3v - dist * dist) / (2.0 * l2 * l3v)
        Dim alpha As Double = Deg(Acos(Clamp(ca)))
        Dim beta As Double = Deg(Acos(Clamp(cb)))

        ' Elbow-up branch: the upper arm stands more upright than the line to the wrist.
        Dim a2abs As Double = toWrist - alpha
        Dim a3abs As Double = a2abs + (180.0 - beta)

        Dim j2 As Double = a2abs
        Dim j3 As Double = (a3abs - a2abs) - J3Zero(p)

        ' The flange axis is the d4 direction, which leads the J3-J5 line by phi3.
        ' Whatever is left over to 180 degrees is what aims the tool at the floor.
        Dim j5 As Double = 180.0 - (a3abs + Phi3(p))

        Dim pose As New RobotPose(name, j1, j2, j3, 0.0, j5, 0.0, "")
        Dim over As String = TravelProblem(pose)
        If over <> "" Then Return New RobotPose(name, j1, j2, j3, 0.0, j5, 0.0, over)
        Return pose
    End Function


    ' All eight key frames, in motion order. Poses repeat where the path does.
    Public Function Sequence(ByVal p As PathInputs) As List(Of RobotPose)
        Dim out As New List(Of RobotPose)
        For Each f As String In Frames
            out.Add(Solve(p, f))
        Next
        Return out
    End Function


    ' Every pose the arm actually passes through, key frames AND the samples that make the
    ' straight legs straight - in motion order, first pose included.
    '
    ' Two things need exactly this list. The driver plays it back; the checks panel tests
    ' it for reach, which the key frames alone cannot do: a straight line between two
    ' reachable points may still leave the workspace or pass over the shoulder in the
    ' middle, and under MoveL the arm really does go there.
    Public Function SampledPath(ByVal p As PathInputs,
                                ByVal tolMm As Double) As List(Of RobotPose)
        Dim out As New List(Of RobotPose)
        out.Add(Solve(p, Frames(0)))

        For k As Integer = 0 To Frames.Length - 2
            Dim a As Double() = Point(p, Frames(k))
            Dim b As Double() = Point(p, Frames(k + 1))
            Dim leg As String = Frames(k) & "->" & Frames(k + 1)

            If Not LegIsLinear(k) Then
                ' MoveJ: the angles interpolate, so there is nothing in between to solve.
                out.Add(Solve(p, Frames(k + 1)))
            Else
                Dim n As Integer = LinearSteps(p, a, b, tolMm)
                For i As Integer = 1 To n
                    Dim s As Double = CDbl(i) / CDbl(n)
                    ' The last sample is the frame itself, so it keeps the frame's name and
                    ' lands on it exactly rather than within rounding of it.
                    If i = n Then
                        out.Add(Solve(p, Frames(k + 1)))
                    Else
                        out.Add(SolveAt(p, leg & " @" & CInt(s * 100) & "%", Lerp(a, b, s)))
                    End If
                Next
            End If
        Next
        Return out
    End Function


    ' Where the tool tip actually lands for a set of angles. Used by the checks panel to
    ' state the round-trip error out loud rather than assert correctness silently.
    '
    ' The two-link chain ends at the WRIST CENTRE (J5), so the tip is tool_len + rb_d6
    ' below it - the same sum SolveAt adds on the way in. It subtracted only tool_len
    ' until 30.08.2026, and the round trip therefore missed by exactly RB_D6: the checks
    ' panel read "Kinematics disagree with themselves by 80 mm" on a model that was right.
    ' The 80 that got fixed in SolveAt, 11_ik_laws and 13_command_laws left this one copy
    ' behind, which is what a check exists to catch - it did.
    Public Function Forward(ByVal pose As RobotPose, ByVal p As PathInputs) As Double()
        Dim a2 As Double = Rad(pose.J2)
        Dim a3 As Double = Rad(pose.J2 + pose.J3 + J3Zero(p))

        Dim r As Double = RB_A1 + p.LenJ2J3 * Sin(a2) + L3(p) * Sin(a3)
        Dim z As Double = RB_D1 + p.LenJ2J3 * Cos(a2) + L3(p) * Cos(a3) -
                          (p.ToolLen + RB_D6)

        Dim j1 As Double = Rad(pose.J1)
        Return New Double() {r * Cos(j1), r * Sin(j1), z}
    End Function


    ' --- limits ----------------------------------------------------------------------------

    Public Function TravelProblem(ByVal pose As RobotPose) As String
        Dim v As Double() = {pose.J1, pose.J2, pose.J3, pose.J4, pose.J5, pose.J6}
        For i As Integer = 0 To 5
            If v(i) < JointMin(i) OrElse v(i) > JointMax(i) Then
                Return "J" & (i + 1) & " = " & Fmt(v(i)) & " deg is outside " &
                       Fmt(JointMin(i)) & " .. " & Fmt(JointMax(i))
            End If
        Next
        Return ""
    End Function


    ' --- helpers -----------------------------------------------------------------------------

    Private Function Bad(ByVal name As String, ByVal why As String) As RobotPose
        Return New RobotPose(name, 0, 0, 0, 0, 0, 0, why)
    End Function

    Private Function Clamp(ByVal v As Double) As Double
        If v > 1.0 Then Return 1.0
        If v < -1.0 Then Return -1.0
        Return v
    End Function

    Private Function Deg(ByVal rad As Double) As Double
        Return rad * 180.0 / PI
    End Function

    Private Function Rad(ByVal deg As Double) As Double
        Return deg * PI / 180.0
    End Function

    Private Function Fmt(ByVal v As Double) As String
        Return v.ToString("0.###", Globalization.CultureInfo.InvariantCulture)
    End Function

End Module
