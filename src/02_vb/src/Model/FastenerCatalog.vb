Option Explicit On
Option Strict Off

' Which stock screw each bolted joint of the machine needs.
'
' FOUR joints, TWO standards, TWO arithmetics:
'
'   post    worktop     -> tapped core hole in the vertical profile's end face (no nut)
'   upper   worktop     -> TNUT_10_M8 in the upper rail's slot
'   lower   shelf       -> TNUT_10_M8 in the lower rail's slot
'   robot   robot base  -> THROUGH the worktop, washer and hex nut underneath
'
' The first three are DIN 7991 countersunk. Their length is measured OVER THE HEAD (the
' head sinks into the plate's countersink), so the arithmetic is literally
' L = thickness + reach, with no head correction - that is the whole reason this screw was
' picked over a cap head. Stock lengths are a coarse series, so L is snapped UP to the next
' one that exists; the surplus goes into the slot, never under the head. That is why every
' joint has a window [MinReach, MaxReach] instead of one number - the tip may go deeper,
' but not so deep it bottoms out on the slot floor or on the end of the tapped core hole.
'
' The fourth is DIN 912 cap head, and there the head sits ON the base flange: nothing of
' the length disappears anywhere, so the number to reach is the whole stack (flange, plate,
' washer, nut, a bit of thread). It still fits the same shape - a fixed amount added to the
' plate thickness - which is why it is a fourth joint here and not a second module. What it
' does NOT have is a hard ceiling: nothing is under the nut for the tip to hit, so MaxReach
' only marks the point where the thread showing below the nut stops looking like a design
' and starts looking like an accident.
'
' Read-only mirror of Rule R_Fasteners in master.CATPart, exactly as ProfileCatalog mirrors
' R_Profile: the rule owns screw_len_post / _upper / _lower / _robot_mount, the app only
' recomputes them so the panel can warn BEFORE anything is sent to CATIA. The rule takes the
' first length that reaches far enough and stops there - a rule cannot refuse to produce
' geometry - so a tip past MaxReach is reported here as a warning, not as a gap.

' One standard in one thread size: the stock lengths it comes in, and the three spellings
' of a length that the rest of the program needs.
'
' Those spellings are not decoration. The DOCUMENT name is how the file is found on disk,
' the PART NUMBER is what the assembly carries and the BOM prints, and the TOKEN is the
' short "M8x30" that both of them contain - which is what makes a swap a string replacement
' in a resolved reference path instead of a second lookup table.
Public Class FastenerFamily

    Public ReadOnly Standard As String        ' "DIN 7991"
    Public ReadOnly Thread As String          ' "M8"
    Private ReadOnly _lengths As Double()
    ' Some suppliers' documents carry the pitch in the file name (DIN_912_M10x50_p1-5).
    ' It is part of the FILE name only - never of the token, so it cannot get in the way of
    ' matching a part number.
    Private ReadOnly _fileSuffix As String

    Public Sub New(ByVal standard_ As String, ByVal thread_ As String,
                   ByVal lengths_ As Double(), ByVal fileSuffix_ As String)
        Standard = standard_
        Thread = thread_
        _lengths = lengths_
        _fileSuffix = fileSuffix_
    End Sub

    Public ReadOnly Property Lengths As Double()
        Get
            Return _lengths
        End Get
    End Property

    Public ReadOnly Property LongestLength As Double
        Get
            Return _lengths(_lengths.Length - 1)
        End Get
    End Property

    ' "M8x30" - the part every spelling of this size has in common.
    Public Function Token(ByVal lengthMm As Double) As String
        Return Thread & "x" & CInt(Math.Round(lengthMm)).ToString()
    End Function

    ' "M8x" - the prefix that recognises ANY length of this family in a part number.
    Public ReadOnly Property TokenPrefix As String
        Get
            Return Thread & "x"
        End Get
    End Property

    ' "DIN 7991 - M8x30", the designation as it goes into the BOM. A standard part is named
    ' by its standard and its size, never by where it happens to sit in this table.
    Public Function Designation(ByVal lengthMm As Double) As String
        Return Standard & " - " & Token(lengthMm)
    End Function

    ' "DIN_7991_M8x30" - the document, without extension.
    Public Function FileStem(ByVal lengthMm As Double) As String
        Return Standard.Replace(" ", "_") & "_" & Token(lengthMm) & _fileSuffix
    End Function

End Class

Public Class FastenerJoint

    Public ReadOnly Key As String            ' post / upper / lower / robot
    Public ReadOnly Caption As String
    Public ReadOnly Family As FastenerFamily
    Public ReadOnly UsesTNut As Boolean
    ' Distance past the underside of the plate, in mm.
    Public ReadOnly MinReach As Double       ' shorter than this and the joint does not hold
    Public ReadOnly MaxReach As Double       ' deeper than this and the tip hits something

    Public Sub New(ByVal key_ As String, ByVal caption_ As String, ByVal family_ As FastenerFamily,
                   ByVal usesTNut_ As Boolean, ByVal minReach_ As Double, ByVal maxReach_ As Double)
        Key = key_
        Caption = caption_
        Family = family_
        UsesTNut = usesTNut_
        MinReach = minReach_
        MaxReach = maxReach_
    End Sub

End Class

' What one joint needs at the current plate thickness.
Public Class FastenerChoice

    Public ReadOnly Joint As FastenerJoint
    Public ReadOnly Thickness As Double
    Public ReadOnly Length As Double         ' chosen stock length, 0 when none fits
    Public ReadOnly Reach As Double          ' Length - Thickness, what actually goes in

    Public Sub New(ByVal joint_ As FastenerJoint, ByVal thickness_ As Double, ByVal length_ As Double)
        Joint = joint_
        Thickness = thickness_
        Length = length_
        Reach = length_ - thickness_
    End Sub

    ' A length exists AND its tip stops before it hits anything.
    Public ReadOnly Property Ok As Boolean
        Get
            Return Length > 0 AndAlso Reach <= Joint.MaxReach
        End Get
    End Property

    ' Long enough to hold, but the next stock size up overshoots the room below the plate.
    Public ReadOnly Property TooDeep As Boolean
        Get
            Return Length > 0 AndAlso Reach > Joint.MaxReach
        End Get
    End Property

    Public ReadOnly Property Designation As String
        Get
            If Length <= 0 Then Return "-"
            Return Joint.Family.Designation(Length)
        End Get
    End Property

End Class

Public Module FastenerCatalog

    ' --- the two families ---------------------------------------------------------------
    '
    ' Stock lengths only. Anything between two of these has to be ordered specially, so the
    ' app never asks for one. The M10 series is coarse (5 mm steps) and starts at 50,
    ' because the shortest stack this joint can have - 15 mm of flange, a thin plate, a
    ' washer, a nut and some thread - already needs about that much.
    Public ReadOnly M8Countersunk As New FastenerFamily(
        "DIN 7991", "M8", New Double() {10, 12, 16, 20, 25, 30, 35, 40, 50, 60}, "")

    Public ReadOnly M10CapHead As New FastenerFamily(
        "DIN 912", "M10", New Double() {50, 55, 60, 65, 70}, "_p1-5")

    ' The two pieces of hardware under the plate. Fixed quantities, fixed sizes: they do not
    ' follow the plate thickness, which is exactly why the app never swaps them - and why
    ' they may stay in a Reuse Pattern while the bolts may not.
    Public Const WasherM10 As String = "DIN 125A - A10,5"
    Public Const NutM10 As String = "DIN 934 - M10"
    Public Const RobotMountCount As Integer = 4

    ' --- DIN 7991 M8, the plate screws ---------------------------------------------------

    ' Head: 90 degree cone, dk = 16.4, k = 4.4. The countersink has to swallow the head
    ' whole, so a plate thinner than the head cannot take this screw at all.
    Public Const HeadHeight As Double = 4.4
    Public Const HeadDiameter As Double = 16.4
    Public Const ShankDiameter As Double = 8.0

    ' Thinnest plate that still has material left under the countersink. One millimetre of
    ' remaining wall is the practical floor; below that the cone breaks through. This is a
    ' property of the COUNTERSINK, so it applies to the three DIN 7991 joints only - the
    ' robot bolt has a flat head on a flat flange and does not care.
    Public Const MinPlateThickness As Double = HeadHeight + 1.0

    ' The slot, measured 16.08.2026 with the nut pushed up against the lip. Everything the
    ' screw meets below the plate, in order:
    '
    '     0.0 .. 3.0    air - the lip, no thread to grip
    '     3.0 .. 11.7   the nut, 8.7 mm of thread
    '    11.7 .. 13.0   clear, then the bottom of the slot
    '
    ' The same on every section, because all five are slot 10 - which is why these are
    ' constants here and an offset in the master, not formulas.
    Public Const TNutAirGap As Double = 3.0
    Public Const TNutHeight As Double = 8.7
    Public Const TNutClearBelow As Double = 1.3

    ' 13 mm is a HARD limit, not a preference: past it the tip lands on the bottom of the
    ' slot and the joint stops clamping, however tight the screw feels.
    Public Const TNutMaxReach As Double = TNutAirGap + TNutHeight + TNutClearBelow

    ' Full engagement would need 11.7 mm of the 13 available, which almost no stock length
    ' hits. 6 mm of thread in a steel nut carries this table with room to spare, and it
    ' opens the window wide enough that only 32-36 mm plates are left unbuildable. Raising
    ' it to 7 kills every fifth thickness (11, 16, 21, 26, ...), which is the whole reason
    ' the number is written down here instead of being picked per joint.
    Public Const TNutMinEngage As Double = 6.0
    Public Const TNutMinReach As Double = TNutAirGap + TNutMinEngage

    ' The post is a plain tapped core hole, so it has room the T-nut does not.
    Public Const PostMinReach As Double = 12.0     ' 1.5 x d, the floor for a thread in aluminium
    Public Const PostMaxReach As Double = 20.0     ' tapped depth of the core hole (VERIFY)

    ' --- DIN 912 M10, the robot base bolts -----------------------------------------------
    '
    ' Everything the bolt passes through that is NOT the worktop. Mirrors the four master
    ' constants rb_base_flange_t / washer_t_m10 / nut_h_m10 / bolt_stickout_min that feed
    ' F_screw_need_robot_mount - the validator compares the two sets against each other.
    Public Const RobotFlangeThickness As Double = 15.0   ' rb_base Pad.2, the mounting flange
    Public Const WasherThicknessM10 As Double = 2.0      ' DIN 125A A10,5
    Public Const NutHeightM10 As Double = 8.0            ' DIN 934 M10
    Public Const StickoutMin As Double = 3.0             ' >= 2 full threads past the nut (P = 1.5)

    ' What the bolt has to cross before any thread shows at all: flange, washer, nut. The
    ' worktop is deliberately not in it - that is the variable the whole exercise turns on.
    Public Const RobotMountGrip As Double = RobotFlangeThickness + WasherThicknessM10 +
                                            NutHeightM10

    ' Read as "what is added to the plate thickness", the same way MinReach is read for the
    ' three countersunk joints - which is what lets Choose() stay one function.
    Public Const RobotMountAdd As Double = RobotMountGrip + StickoutMin

    ' No hard ceiling exists under the nut, so this one is a matter of tidiness: more than
    ' 15 mm of thread standing below the worktop is a bolt that should have been shorter.
    Public Const StickoutMax As Double = 15.0

    Public ReadOnly Post As New FastenerJoint(
        "post", "Worktop to post (tapped core hole)", M8Countersunk, False,
        PostMinReach, PostMaxReach)

    Public ReadOnly UpperPlate As New FastenerJoint(
        "upper", "Worktop to upper rails (T-nut)", M8Countersunk, True,
        TNutMinReach, TNutMaxReach)

    Public ReadOnly LowerPlate As New FastenerJoint(
        "lower", "Shelf to lower rails (T-nut)", M8Countersunk, True,
        TNutMinReach, TNutMaxReach)

    Public ReadOnly RobotMount As New FastenerJoint(
        "robot", "Robot base to worktop (through-bolt, washer + nut)", M10CapHead, False,
        RobotMountAdd, RobotMountGrip + StickoutMax)

    ' The shortest stock length that reaches far enough. Length 0 means the series runs
    ' out - the plate is thicker than this family can fasten. Whether the tip then goes too
    ' deep is a separate question, asked through FastenerChoice.TooDeep, because the rule in
    ' the master reaches the same answer and has no way to say "none".
    Public Function Choose(ByVal joint As FastenerJoint, ByVal thickness As Double) As FastenerChoice
        Dim lo As Double = thickness + joint.MinReach
        For Each l As Double In joint.Family.Lengths
            If l >= lo Then Return New FastenerChoice(joint, thickness, l)
        Next
        Return New FastenerChoice(joint, thickness, 0)
    End Function

    ' Thickest plate this joint can still fasten: the longest screw, minus the shortest
    ' reach that still holds. Above it no stock length is long enough.
    Public Function MaxThickness(ByVal joint As FastenerJoint) As Double
        Return joint.Family.LongestLength - joint.MinReach
    End Function

    ' Every screw the machine needs, per joint. The robot bolt goes through the same plate
    ' as the first two, so it moves with the worktop thickness the way they do.
    Public Function Plan(ByVal d As TableDesign) As List(Of FastenerChoice)
        Dim list As New List(Of FastenerChoice)()
        list.Add(Choose(Post, d.UpperPlateThickness))
        list.Add(Choose(UpperPlate, d.UpperPlateThickness))
        list.Add(Choose(LowerPlate, d.LowerPlateThickness))
        list.Add(Choose(RobotMount, d.UpperPlateThickness))
        Return list
    End Function

    ' One length per JOINT, not per plate. Sharing a length across the worktop's two M8
    ' joints was measured on 16.08.2026 to be impossible: at 15 mm the post wants M8 x 30
    ' and the T-nut wants M8 x 25, and the 30 would bottom out in the slot (15 mm of reach
    ' against 13 available). The tapped core hole simply has room the slot has not.

End Module
