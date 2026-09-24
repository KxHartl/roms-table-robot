Option Explicit On
Option Strict Off

' Base mounting data of the robot this table carries.
'
' ONE ROBOT ON PURPOSE (decision 17.08.2026). The task assigns the FANUC LR Mate 200iC and
' that is the only robot being modelled, so a picker with a list of robots would offer
' choices that lead nowhere - no CAD, no gripper, no kinematics behind them. What the
' window needs is a readout of THIS robot next to the raster fields, so the numbers being
' written are visibly somebody's data and not a guess.
'
' The raster fields stay editable regardless: task section 3 asks for the bolt raster to be
' changeable, so this record is a starting point, never a lock.
'
' ZERO MEANS UNKNOWN, NOT ZERO. A value nobody has read off a manufacturer drawing is 0
' and `Verified` is False, and the UI says so. A made-up bolt raster is worse than no
' raster: it drills four holes in the wrong place and nothing in the model can catch it.
'
' The same data, with sources, is in src/02_vb/robots.txt. Change one, change both.
Public Class RobotInfo

    Public ReadOnly Name As String            ' key, same spelling as the RoboDK library
    Public ReadOnly Caption As String         ' what the readout shows

    ' Base bolt raster. Long = the wider pitch, short = the narrower one; which of them
    ' lands on X depends on how the robot is turned (TableDesign.MountPitchX).
    Public ReadOnly BoltPitchLong As Double   ' mm, 0 = unknown
    Public ReadOnly BoltPitchShort As Double  ' mm, 0 = unknown
    Public ReadOnly HoleDiameter As Double    ' mm, clearance hole in the worktop, 0 = unknown
    Public ReadOnly BoltThread As String      ' "M8", "" = unknown

    Public ReadOnly Reach As Double           ' mm, 0 = unknown
    Public ReadOnly Payload As Double         ' kg, 0 = unknown
    Public ReadOnly LenJ2J3 As Double         ' mm, 0 = unknown
    Public ReadOnly LenJ3J5 As Double         ' mm, 0 = unknown

    Public ReadOnly Verified As Boolean       ' read off a manufacturer drawing?
    Public ReadOnly Source As String

    Public Sub New(ByVal name_ As String, ByVal caption_ As String,
                   ByVal pitchLong As Double, ByVal pitchShort As Double,
                   ByVal holeD As Double, ByVal thread As String,
                   ByVal reach_ As Double, ByVal payload_ As Double,
                   ByVal j2j3 As Double, ByVal j3j5 As Double,
                   ByVal verified_ As Boolean, ByVal source_ As String)
        Name = name_
        Caption = caption_
        BoltPitchLong = pitchLong
        BoltPitchShort = pitchShort
        HoleDiameter = holeD
        BoltThread = thread
        Reach = reach_
        Payload = payload_
        LenJ2J3 = j2j3
        LenJ3J5 = j3j5
        Verified = verified_
        Source = source_
    End Sub

    ' True when the three values the app would write are all present.
    Public ReadOnly Property HasRaster As Boolean
        Get
            Return BoltPitchLong > 0 AndAlso BoltPitchShort > 0 AndAlso HoleDiameter > 0
        End Get
    End Property

End Class

Public Module RobotCatalog

    ' The robot the task assigns, and the only one modelled.
    '
    ' Raster (205 x 160, M10) confirmed 18.08.2026 by the user modelling the base by hand
    ' against src/01_cad/01_mainAssembly/LR Mate-200iC base.CATPart - replaces the earlier
    ' 100 x 60 / M8 guess. HoleDiameter = 11: decided 27.08.2026 - the M10 bolt is
    ' clearance-fit through the plate with a nut underneath (ISO medium clearance, same
    ' diameter as the robot's own through-hole); robot_mount_hole_d = 11 is set in the
    ' master. The robot's OWN hole (rb_base_hole_d, local to rb_base.CATPart) is the same
    ' number by coincidence, not by link. Reach/payload from the FANUC LR Mate 200iC
    ' datasheet; J2-J3 / J3-J5 nominal from the RoboDK model.
    Public ReadOnly Assigned As New RobotInfo(
        "Fanuc-LR-Mate-200iC", "FANUC LR Mate 200iC",
        205.0, 160.0, 11.0, "M10",
        704.0, 5.0, 300.0, 320.0,
        True,
        "205 x 160 / M10 raster from the vendor STEP; plate hole dia 11 (M10 clearance + nut). " &
        "Reach/payload: FANUC datasheet. J2-J3 / J3-J5: RoboDK. Full provenance in robots.txt.")

End Module
