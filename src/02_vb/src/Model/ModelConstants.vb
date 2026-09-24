Option Explicit On
Option Strict Off

' Constants the model owns and the app never writes (PARAMETERS.md section 2), plus the
' robot values that come from RoboDK. They are READ from master.CATPart on connect, so
' the limits follow the model instead of a second hard-coded truth. The defaults below
' are only used before the first connection.
Public Class ModelConstants

    Public FeetHeight As Double = 40.0            ' feet_height
    Public PlateClearance As Double = 0.5         ' plate_clearance, per side
    Public HoleEdgeDistance As Double = 30.0      ' hole_edge_dist: plate edge -> hole axis

    ' MOVED OUT 17.08.2026: robot_mount_length / robot_mount_width are no longer constants.
    ' Task section 3 wants the robot bolt raster changeable from the app, so they became
    ' inputs (TableDesign.MountPitchLong / MountPitchShort) and the app writes them. The
    ' hole diameter went the same way, from an app-only drawing number to the parameter
    ' robot_mount_hole_d. Do not reintroduce them here - two owners for one number is
    ' exactly what PARAMETERS.md forbids.

    ' Nominal segment lengths read from RoboDK; the task allows nominal .. nominal + 20 %.
    Public NominalLenJ2J3 As Double = 300.0
    Public NominalLenJ3J5 As Double = 320.0

    ' False makes the app treat the base offsets as if the hole sketch still applied them
    ' rotated by 90 degrees. See PARAMETERS.md -> R_RobotMountOffset.
    Public MountSketchFixed As Boolean = True

    Public Sub New()
        NominalLenJ2J3 = AppConfig.NominalLenJ2J3()
        NominalLenJ3J5 = AppConfig.NominalLenJ3J5()
        MountSketchFixed = AppConfig.RobotMountSketchFixed()
    End Sub

    ' Overwrite the defaults with whatever the master part actually holds. Missing
    ' parameters keep the default - the app must stay usable against an older model.
    Public Sub ReadFrom(ByVal params As MasterParameters)
        If params Is Nothing Then Return
        FeetHeight = Pick(params, "feet_height", FeetHeight)
        PlateClearance = Pick(params, "plate_clearance", PlateClearance)
        HoleEdgeDistance = Pick(params, "hole_edge_dist", HoleEdgeDistance)
    End Sub

    Private Shared Function Pick(ByVal params As MasterParameters, ByVal name As String, ByVal dflt As Double) As Double
        Dim raw As String = params.ReadNumber(name)
        Dim d As Double
        If Double.TryParse(Num.Normalize(raw), Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, d) Then
            If d > 0 Then Return d
        End If
        Return dflt
    End Function

End Class
