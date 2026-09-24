Option Explicit On
Option Strict Off

' The two grippers the task assigns, and the objects they handle.
'
' Same contract as RobotCatalog: ZERO MEANS UNKNOWN, NOT ZERO. A number nobody has read
' off a drawing or measured in the CAD is 0 and Verified is False, and the window says
' so. A made-up bolt spacing drills two holes in the wrong place on the adapter plate,
' and nothing in the model can catch that - the plate simply will not fit a gripper that
' exists only in a supplier's warehouse.
'
' This mirrors the Knowledge rule R_Tool, the way ProfileCatalog mirrors R_Profile and
' FastenerCatalog mirrors R_Fasteners. The rule owns the model; this owns what the
' window can say before anything is written. Change one, change the other -
' 02_validate_master compares them.
'
' The folders used to be swapped (EGP 64-N-N-IOL/ held MPG-plus 40 and vice versa).
' Renamed 28.08.2026 (TODO.md B1), so FolderToday below now matches the part number:
'     src/01_cad/07_schunk/MPG-plus-40/     -> MPG-plus 40  (0305521)
'     src/01_cad/07_schunk/EGP 64-N-N-IOL/  -> EGP 64-N-N-IOL (1383545)

Public Class GripperInfo

    Public ReadOnly Key As String              ' object_type value: "1" or "2"
    Public ReadOnly Caption As String
    Public ReadOnly OrderNo As String

    ' --- what the object is --------------------------------------------------------------
    Public ReadOnly ObjA As Double              ' width  [mm]
    Public ReadOnly ObjB As Double              ' length [mm]
    Public ReadOnly ObjC As Double              ' height [mm]

    ' --- what the gripper can do ----------------------------------------------------------
    Public ReadOnly StrokePerJaw As Double      ' mm, 0 = unknown
    Public ReadOnly BodyHeight As Double        ' mounting face to jaw face [mm]
    Public ReadOnly MaxFingerLen As Double      ' catalogue limit [mm]
    Public ReadOnly MaxFingerMass As Double     ' kg per finger
    Public ReadOnly Mass As Double              ' kg, for the BOM

    ' --- how it bolts to the adapter plate -------------------------------------------------
    Public ReadOnly MountThread As String       ' "M4", "M5"
    Public ReadOnly MountPitch As Double        ' mm between the two mounting threads
    Public ReadOnly MountSleeveD As Double      ' centring sleeve diameter [mm]

    ' --- how the fingers bolt to the base jaws ---------------------------------------------
    Public ReadOnly FingerThread As String
    Public ReadOnly FingerPinD As Double        ' centring pin diameter [mm]
    Public ReadOnly FingerSeatW As Double       ' locating step width [mm]

    Public ReadOnly FolderToday As String       ' repo folder that holds this gripper's CAD
    Public ReadOnly Verified As Boolean
    Public ReadOnly Source As String

    Public Sub New(ByVal key_ As String, ByVal caption_ As String, ByVal orderNo_ As String,
                   ByVal a As Double, ByVal b As Double, ByVal c As Double,
                   ByVal stroke As Double, ByVal bodyH As Double,
                   ByVal maxFingerLen_ As Double, ByVal maxFingerMass_ As Double,
                   ByVal mass_ As Double,
                   ByVal mountThread_ As String, ByVal mountPitch_ As Double,
                   ByVal mountSleeveD_ As Double,
                   ByVal fingerThread_ As String, ByVal fingerPinD_ As Double,
                   ByVal fingerSeatW_ As Double,
                   ByVal folderToday_ As String,
                   ByVal verified_ As Boolean, ByVal source_ As String)
        Key = key_ : Caption = caption_ : OrderNo = orderNo_
        ObjA = a : ObjB = b : ObjC = c
        StrokePerJaw = stroke : BodyHeight = bodyH
        MaxFingerLen = maxFingerLen_ : MaxFingerMass = maxFingerMass_ : Mass = mass_
        MountThread = mountThread_ : MountPitch = mountPitch_ : MountSleeveD = mountSleeveD_
        FingerThread = fingerThread_ : FingerPinD = fingerPinD_ : FingerSeatW = fingerSeatW_
        FolderToday = folderToday_
        Verified = verified_ : Source = source_
    End Sub

    ' Total jaw travel. The fingers have to bridge everything between this and the
    ' object - which is the entire reason the task asks for custom fingers.
    Public ReadOnly Property TotalStroke As Double
        Get
            Return StrokePerJaw * 2.0
        End Get
    End Property

End Class


Public Module GripperCatalog

    ' Object 1: 40 x 50 x 60, handled by the MPG-plus 40.
    Public ReadOnly Object1 As New GripperInfo(
        "1", "SCHUNK MPG-plus 40", "0305521",
        40.0, 50.0, 60.0,
        6.0, 39.0, 50.0, 0.08, 0.18,
        "M4", 32.0, 6.0,
        "M5", 2.5, 16.0,
        "src/01_cad/07_schunk/MPG-plus-40",
        True,
        "SCHUNK MPG-plus 40 (stroke 6 mm per jaw, body 40 x 26 x 39). Measured 28.08.2026 " &
        "in the CAD: mount 2 x M4 at 32 with a 6 mm centring sleeve; finger M5 with a " &
        "2.5 mm pin offset ~3 / ~5 mm from the thread, on a 16 (0/-0.04) side seat. " &
        "gripper_h / finger_reach may still shift when the adapter and finger are modelled.")

    ' Object 2: 80 x 80 x 80, handled by the EGP 64-N-N-IOL.
    Public ReadOnly Object2 As New GripperInfo(
        "2", "SCHUNK EGP 64-N-N-IOL", "1383545",
        80.0, 80.0, 80.0,
        10.0, 115.0, 80.0, 0.24, 0.83,
        "M5", 52.0, 8.0,
        "M8", 4.0, 20.0,
        "src/01_cad/07_schunk/EGP 64-N-N-IOL",
        True,
        "SCHUNK EGP 64-N-N-IOL (stroke 10 mm per jaw - HUBBACKE = 10.000 inside the " &
        ".catproduct; body 64 x 35 x 114.7). Measured 28.08.2026: mount 2 x M5 at 52 with " &
        "an 8 mm centring sleeve; finger M8 with a 4 mm pin offset ~4.5 / ~6 mm from the " &
        "thread, on a 20 (0/-0.04) side seat. gripper_h / finger_reach may still shift.")

    ' Flange to gripping point, MEASURED in the assembly on 30.08.2026: TOOL_AXS - which
    ' is coincident with the flange's FLANGE_AXS - to TCP_AXS.
    '
    ' NOT adapter_t + BodyHeight + MaxFingerLen, and not any other sum of catalogue
    ' figures. That chain was the model's for a while and it is wrong in both
    ' configurations: it gives 96 where the tool measures 122, and 197 where it measures
    ' 221. The parts are shaped to technical need and owe the arithmetic nothing.
    '
    ' The authority is the rule R_Tool inside the master, which carries the same two
    ' numbers. This copy exists so the window can draw a path and check reach with no
    ' CATIA running; TableDesign prefers the master's value whenever it has one, and the
    ' checks panel says which of the two it used.
    Public Function MeasuredToolLen(ByVal objectType As String) As Double
        If Trim(If(objectType, "")) = "2" Then Return 221.0
        Return 122.0
    End Function


    Public Function ByType(ByVal objectType As String) As GripperInfo
        If objectType IsNot Nothing AndAlso objectType.Trim() = "2" Then Return Object2
        Return Object1
    End Function

    Public Function All() As GripperInfo()
        Return New GripperInfo() {Object1, Object2}
    End Function

    ' Can this gripper hold this object at all, given the fingers may only bridge the
    ' gap and not travel themselves? The jaws move TotalStroke between fully open and
    ' fully closed; the fingers are what set where that window sits. So the question is
    ' not "is the object smaller than the stroke" - it is whether a finger can be built
    ' whose grip range brackets the object, and whether that finger is inside the
    ' catalogue's length limit.
    '
    ' Returns "" when it is fine, otherwise the reason to show the user.
    Public Function GripProblem(ByVal g As GripperInfo, ByVal fingerReach As Double) As String
        If g Is Nothing Then Return "no gripper selected"

        If g.StrokePerJaw <= 0 Then
            Return "stroke of the " & g.Caption & " is unknown - measure it before trusting this"
        End If

        If fingerReach > g.MaxFingerLen AndAlso g.MaxFingerLen > 0 Then
            Return "finger reach " & Num.MmString(fingerReach) & " exceeds the catalogue " &
                   "limit of " & Num.MmString(g.MaxFingerLen) & " for the " & g.Caption &
                   " - the jaws are not rated for the bending moment"
        End If

        ' A grip needs the object to be releasable: the jaws must open at least a
        ' visible amount past the object, or the tool cannot be withdrawn.
        If g.TotalStroke < 2.0 Then
            Return "total jaw travel of " & Num.MmString(g.TotalStroke) &
                   " leaves no clearance to release the object"
        End If

        Return ""
    End Function

End Module
