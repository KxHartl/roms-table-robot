Option Explicit On
Option Strict Off

' The values currently shown in the UI plus everything the model derives from them.
'
' The derived properties are a faithful copy of the formulas in PARAMETERS.md (sections
' 3-5). They exist so the UI can pre-compute limits, draw the sketch and warn BEFORE
' anything reaches CATIA. Nothing here is ever written to the model - rules and formulas
' keep sole ownership of these parameters.
Public Class TableDesign

    ' --- user inputs (the app writes exactly these) ---
    Public Length As Double = 550.0
    Public Width As Double = 420.0
    Public Height As Double = 450.0
    Public UpperPlateThickness As Double = 15.0
    Public LowerPlateThickness As Double = 10.0
    Public ProfileName As String = "45x60"
    Public LowerPlateVariant As String = "v2"
    Public RobotMountVariant As String = "centerShortAxis"
    Public Simplified As Boolean = False
    Public UserOffsetX As Double = 20.0
    Public UserOffsetY As Double = 60.0

    ' Robot base bolt raster and hole. Inputs since 17.08.2026 (task section 3 asks for the
    ' raster to be changeable): robot_mount_length / robot_mount_width used to be constants
    ' set by hand in CATIA, and the hole diameter existed only in the app as a drawing
    ' number. Long / short rather than X / Y - which one lands on X is the mount variant's
    ' business (MountPitchX below).
    Public MountPitchLong As Double = RobotCatalog.Assigned.BoltPitchLong    ' robot_mount_length
    Public MountPitchShort As Double = RobotCatalog.Assigned.BoltPitchShort  ' robot_mount_width
    Public MountHoleDiameter As Double = RobotCatalog.Assigned.HoleDiameter  ' robot_mount_hole_d

    ' App-side only: how far the robot base body reaches past the outer bolts. It sets the
    ' minimum distance from the plate edge and the footprint drawn in the sketch, and is
    ' never written to the master (the model has no such parameter).
    Public BaseOverhang As Double = AppConfig.RobotBaseOverhang()

    ' --- robot and path inputs ---
    ' These stopped being "optional" once 09_robot_parameters started creating them.
    ' They are still written through TryOptional, so an older master without them is
    ' skipped rather than failed - but the defaults below are now the model's defaults,
    ' not placeholders.
    Public LenJ2J3 As Double = 300.0
    Public LenJ3J5 As Double = 320.0

    ' P1, P3 and P6 are in the ROBOT BASE FRAME: origin on the robot's mounting face, X
    ' forward, Z up. Not the table frame and not the floor. The inverse kinematics works
    ' in the robot's own frame, and a second conversion is how a path ends up half a
    ' metre below where it was meant to be.
    Public P1X As Double = 275.0
    Public P1Y As Double = 60.0
    Public P1Z As Double = 600.0

    ' The pick station, and where the cube is initialised - its CENTRE sits here, because
    ' that is the point the tool grips. P2, P4 and P5 follow from it; only P4P5 is derived
    ' (= P2P3), since the task figure puts both stations at one height.
    Public P3X As Double = 412.0
    Public P3Y As Double = -80.0
    Public P3Z As Double = 150.0

    Public P6X As Double = 275.0
    Public P6Y As Double = -60.0
    Public P6Z As Double = 600.0

    ' P2P4 runs along Y, P2P3 along Z. Positive P2P4 moves the place station toward +Y,
    ' which is why P3Y is negative: at P3Y = -80 and P2P4 = 160 the stations land at -80
    ' and +80, symmetric about the robot's X axis.
    Public P2P3 As Double = 60.0
    Public P2P4 As Double = 160.0

    ' --- simulation speed ---
    ' Different for the two motion types, as a robot is actually programmed: fast through
    ' free space, at a set speed in the working moves.
    Public SimLeg As Double = 1.0      ' seconds per MoveJ leg
    Public SimSpeed As Double = 100.0  ' mm/s on the MoveL legs

    ' --- tool ---
    ' The reconfiguration switch of task section 9. R_Tool branches on it; ToolSwitcher
    ' exchanges sub_tool_1 for sub_tool_2.
    Public ObjectType As String = "1"
    ' Adapter plate thickness. Not an input, not written to the master, and no longer part
    ' of any tool-length arithmetic - it is kept only because DesignLimits still reports it
    ' when it explains a finger design.
    Public AdapterT As Double = 12.0

    ' tool_len as the MASTER holds it, and which object type it was read for.
    '
    ' The master is the authority: the rule R_Tool sets tool_len from measured values, 122
    ' for tool 1 and 221 for tool 2. The window reads it on connect and on reload.
    '
    ' The object type is remembered alongside it because the two can disagree. Pick the
    ' other object in the window and the master still holds the length of the tool that is
    ' actually fitted - so the number is only trusted while it belongs to the type now
    ' selected, and the catalogue value stands in for the preview until Apply makes them
    ' agree again.
    Public ToolLenFromModel As Double = 0.0
    Public ToolLenForType As String = ""

    ' The gripper the current ObjectType implies. Read-only convenience for the checks
    ' panel; the authority is the rule R_Tool inside the master.
    Public ReadOnly Property Gripper As GripperInfo
        Get
            Return GripperCatalog.ByType(ObjectType)
        End Get
    End Property

    ' Everything RobotKinematics needs, in its own shape. Keeping the conversion in one
    ' place means the window, the checks and the simulate button can never disagree
    ' about which numbers the arm is being asked to reach.
    Public Function ToPathInputs() As RobotKinematics.PathInputs
        Dim p As New RobotKinematics.PathInputs()
        p.LenJ2J3 = LenJ2J3
        p.LenJ3J5 = LenJ3J5
        p.P1X = P1X
        p.P1Y = P1Y
        p.P1Z = P1Z
        p.P3X = P3X
        p.P3Y = P3Y
        p.P3Z = P3Z
        p.P6X = P6X
        p.P6Y = P6Y
        p.P6Z = P6Z
        p.P2P3 = P2P3
        p.P2P4 = P2P4
        p.ToolLen = ToolLen()
        p.GripZ = Gripper.ObjC / 2.0
        Return p
    End Function

    ' Flange to gripping point. The master's measured value when it applies, the catalogue
    ' measurement otherwise.
    '
    ' This used to be AdapterT + BodyHeight + FingerReach - a sum of catalogue figures that
    ' the finished parts never owed anything to. It reads 96 for tool 1 and 197 for tool 2;
    ' measured, they are 122 and 221. The inverse kinematics places the wrist exactly this
    ' far above the target, so the error went straight into the tool tip's height and the
    ' arm reached for the cube some hundred millimetres off.
    Public Function ToolLen() As Double
        If ToolLenFromModel > 0.0 AndAlso ToolLenForType = ObjectType Then
            Return ToolLenFromModel
        End If
        Return GripperCatalog.MeasuredToolLen(ObjectType)
    End Function


    ' Which of the two the number above came from, for the checks panel. Two sources that
    ' silently disagree is the fault this whole change exists to remove, so the window says
    ' out loud which one it is using.
    Public Function ToolLenSource() As String
        If ToolLenFromModel > 0.0 AndAlso ToolLenForType = ObjectType Then
            Return "master"
        End If
        If ToolLenFromModel > 0.0 Then
            Return "catalogue - the master still holds the fitted tool's length"
        End If
        Return "catalogue - not read from the master yet"
    End Function


    ' What the task says the two segments should be for the object now fitted:
    ' P2-P4 = 4a and P2-P3 = c, where a and c are the object's own dimensions.
    '
    ' PROPOSED, not imposed. The window writes these when the object type changes so a run
    ' always starts from what the task asks for, but both stay ordinary inputs the user can
    ' then move - task section 12 asks for them to be changeable, and a formula would make
    ' them read-only. Same contract as GripperCatalog: the catalogue suggests, the user
    ' decides.
    Public Function SuggestedP2P4() As Double
        Return 4.0 * Gripper.ObjA
    End Function

    Public Function SuggestedP2P3() As Double
        Return Gripper.ObjC
    End Function

    ' Jaw face to the gripping point. The model keeps this as finger_reach, driven by
    ' R_Tool; here it follows the same two numbers so the preview matches.
    Public Function FingerReach() As Double
        If ObjectType = "2" Then Return 70.0
        Return 45.0
    End Function

    Public Constants As New ModelConstants()

    Public ReadOnly Property Profile As ProfileInfo
        Get
            Return ProfileCatalog.Find(ProfileName)
        End Get
    End Property

    ' --- section halves (R_Profile) ---
    Public ReadOnly Property OffsetLength As Double  ' along X
        Get
            Return Profile.OffsetLength
        End Get
    End Property

    Public ReadOnly Property OffsetWidth As Double   ' along Y
        Get
            Return Profile.OffsetWidth
        End Get
    End Property

    Public ReadOnly Property GussetSize As Double
        Get
            Return Profile.GussetSize
        End Get
    End Property

    ' The other bracket size: rail-to-rail, one in each corner of each level (8 of them).
    ' Equal to GussetSize on the square sections, larger on 45x60 and 45x90.
    Public ReadOnly Property GussetSizeCorner As Double
        Get
            Return Profile.GussetSizeCorner
        End Get
    End Property

    Public ReadOnly Property GussetAbove As Boolean  ' V2 mounts the brackets above the rail
        Get
            Return Not String.Equals(LowerPlateVariant, "v1", StringComparison.OrdinalIgnoreCase)
        End Get
    End Property

    ' --- table skeleton (PARAMETERS.md section 5) ---
    Public ReadOnly Property PostCenterX As Double
        Get
            Return Length / 2.0 - OffsetLength
        End Get
    End Property

    Public ReadOnly Property PostCenterY As Double
        Get
            Return Width / 2.0 - OffsetWidth
        End Get
    End Property

    Public ReadOnly Property SpanLength As Double
        Get
            Return Length - 4.0 * OffsetLength
        End Get
    End Property

    Public ReadOnly Property SpanWidth As Double
        Get
            Return Width - 4.0 * OffsetWidth
        End Get
    End Property

    Public ReadOnly Property HeightVerticalProfile As Double
        Get
            Return Height - UpperPlateThickness - Constants.FeetHeight
        End Get
    End Property

    ' R_TableVariants: V1 hangs the brackets under the rail, so the whole lower level
    ' climbs by one bracket size.
    Public ReadOnly Property HeightLowerHorizontal As Double
        Get
            Dim h As Double = Constants.FeetHeight + OffsetLength
            If Not GussetAbove Then h += GussetSize
            Return h
        End Get
    End Property

    ' Centreline of the upper rails - mirror of F_height_upper_horizontal. The corner
    ' brackets of the upper level are anchored at exactly this height.
    Public ReadOnly Property HeightUpperHorizontal As Double
        Get
            Return Height - UpperPlateThickness - OffsetLength
        End Get
    End Property

    ' Clear vertical distance between the two rail levels (Formula.31 / .32).
    Public ReadOnly Property RailGap As Double
        Get
            Return HeightUpperHorizontal - HeightLowerHorizontal
        End Get
    End Property

    Public ReadOnly Property LowerPlateLength As Double
        Get
            If GussetAbove Then
                Return SpanLength - 2.0 * GussetSize - 2.0 * Constants.PlateClearance
            End If
            Return SpanLength - 2.0 * Constants.PlateClearance
        End Get
    End Property

    Public ReadOnly Property LowerPlateWidth As Double
        Get
            Return Width
        End Get
    End Property

    Public ReadOnly Property LowerPlateZ As Double
        Get
            Return HeightLowerHorizontal + OffsetLength
        End Get
    End Property

    ' --- robot base raster (R_RobotMountOffset) ---
    ' "corner" is the retired name of cornerLongAxis; models written before the free
    ' placement still carry it, and the rule's else branch still catches it.
    Public ReadOnly Property FreePosition As Boolean
        Get
            Return RobotMountVariant = "cornerLongAxis" OrElse
                   RobotMountVariant = "cornerShortAxis" OrElse
                   RobotMountVariant = "corner"
        End Get
    End Property

    Public ReadOnly Property UsesOffsetX As Boolean
        Get
            Return RobotMountVariant = "edgeLongAxis" OrElse FreePosition
        End Get
    End Property

    Public ReadOnly Property UsesOffsetY As Boolean
        Get
            Return RobotMountVariant = "edgeShortAxis" OrElse FreePosition
        End Get
    End Property

    ' The robot the raster came from. Never written to the model - the master has no
    ' parameter for it, and the raster it drives is written instead.
    Public ReadOnly Property Robot As RobotInfo
        Get
            Return RobotCatalog.Assigned
        End Get
    End Property

    ' False once any of the three raster inputs has been moved away from the robot's own
    ' figures, so the UI can say "these are your numbers now, not the robot's".
    Public ReadOnly Property MatchesCatalogue As Boolean
        Get
            Dim r As RobotInfo = Robot
            If Not r.HasRaster Then Return False
            Return Math.Abs(MountPitchLong - r.BoltPitchLong) < 0.01 AndAlso
                   Math.Abs(MountPitchShort - r.BoltPitchShort) < 0.01 AndAlso
                   Math.Abs(MountHoleDiameter - r.HoleDiameter) < 0.01
        End Get
    End Property

    ' Pitch carries the orientation: pitch_x = the long pitch means the long side of the
    ' base lies along the table length.
    Public ReadOnly Property MountPitchX As Double
        Get
            Select Case RobotMountVariant
                Case "centerShortAxis", "edgeShortAxis", "cornerShortAxis" : Return MountPitchShort
                Case Else : Return MountPitchLong
            End Select
        End Get
    End Property

    Public ReadOnly Property MountPitchY As Double
        Get
            Select Case RobotMountVariant
                Case "centerShortAxis", "edgeShortAxis", "cornerShortAxis" : Return MountPitchLong
                Case Else : Return MountPitchShort
            End Select
        End Get
    End Property

    ' Outer hole, measured from the plate centre.
    Public ReadOnly Property AxisOffsetX As Double
        Get
            If UsesOffsetX Then Return Length / 2.0 - UserOffsetX
            Return MountPitchX / 2.0
        End Get
    End Property

    Public ReadOnly Property AxisOffsetY As Double
        Get
            If UsesOffsetY Then Return Width / 2.0 - UserOffsetY
            Return MountPitchY / 2.0
        End Get
    End Property

    ' Plate edge at +X / +Y to the first bolt. This is what the hole sketch measures,
    ' because a sketch dimension cannot be negative (F_mount_edge_x/y in the master).
    Public ReadOnly Property MountEdgeX As Double
        Get
            Return Length / 2.0 - AxisOffsetX
        End Get
    End Property

    Public ReadOnly Property MountEdgeY As Double
        Get
            Return Width / 2.0 - AxisOffsetY
        End Get
    End Property

    ' The four bolt holes: outer hole plus one pitch inwards, per axis.
    Public Function MountHolesX() As Double()
        Return New Double() {AxisOffsetX, AxisOffsetX - MountPitchX}
    End Function

    Public Function MountHolesY() As Double()
        Return New Double() {AxisOffsetY, AxisOffsetY - MountPitchY}
    End Function

End Class
