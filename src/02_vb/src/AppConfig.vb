Option Explicit On
Option Strict Off

Imports System.Configuration
Imports System.IO

' Centralizirana konfiguracija: katalozi dopustenih vrijednosti + putanje (App.config).
Public Module AppConfig

    ' Katalozi (mirror parameters_desk.xlsx / VBA forme).
    Public ReadOnly Profiles As String() = {"45x45", "45x60", "45x90", "60x60", "90x90"}
    ' Samo v1 i v2. "v3" je bio visak iz stare VBA forme - R_TableVariants ima dvije grane
    ' (kutnici ispod / iznad ukrute), pa bi v3 samo tiho pao u else i ponasao se kao v1.
    Public ReadOnly LowerPlateMounts As String() = {"v1", "v2"}
    ' Six values: three positions (centre / edge / corner) x two orientations
    ' (long side along the length = X, or across the width = Y). R_RobotMountOffset
    ' branches on all six; the retired "corner" still lands in the LongAxis else.
    Public ReadOnly RobotMounts As String() = {"centerLongAxis", "centerShortAxis", "edgeLongAxis", "edgeShortAxis", "cornerLongAxis", "cornerShortAxis"}
    Public ReadOnly Simplifications As String() = {"false", "true"}

    Public Function Setting(ByVal key As String, ByVal dflt As String) As String
        Dim v As String = Nothing
        Try
            v = ConfigurationManager.AppSettings(key)
        Catch
            v = Nothing
        End Try
        If String.IsNullOrEmpty(v) Then Return dflt
        Return v
    End Function

    ' A relative CadRoot (the handed-in package ships "..\03_CATIA_model") is taken from
    ' the folder the .exe sits in, so the package runs wherever it is unpacked. The
    ' absolute default is the development tree.
    Public Function CadRoot() As String
        Dim p As String = Setting("CadRoot", "D:\truenas_kresimir_share_cp\FSB\semestar_07\roms\projektni_zadatak\src\01_cad")
        If Not Path.IsPathRooted(p) Then
            p = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, p))
        End If
        Return p
    End Function

    ' One CAD document, named by the sub-folder it has in the development tree.
    ' CATIA's Send To > Directory - how the model is handed in, because it rewrites the
    ' links - copies every document into ONE flat folder, so when root\subFolder\fileName
    ' is missing the document is looked for directly in root. If neither exists the
    ' structured path comes back, so "not found" still names where it was expected.
    Public Function CadFile(ByVal root As String, ByVal subFolder As String, ByVal fileName As String) As String
        Dim structured As String = Path.Combine(root, subFolder, fileName)
        If File.Exists(structured) Then Return structured
        Dim flat As String = Path.Combine(root, fileName)
        If File.Exists(flat) Then Return flat
        Return structured
    End Function

    Public Function MasterPartPath() As String
        Dim p As String = Setting("MasterPartPath", "")
        If String.IsNullOrEmpty(p) Then p = CadFile(CadRoot(), "00_main", "master.CATPart")
        Return p
    End Function

    Public Function TableAssemblyPath() As String
        Dim p As String = Setting("TableAssemblyPath", "")
        If String.IsNullOrEmpty(p) Then p = CadFile(CadRoot(), "01_table", "table_mainAssembly.CATProduct")
        Return p
    End Function

    ' The robot assembly, where the DMU mechanism and the tool live. Opened lazily -
    ' only when the tool has to be swapped or the path simulated - so someone working
    ' on the table alone never waits for it to load.
    '
    ' It lives in 05_robot, NOT 00_main - that folder holds the master and the
    ' main assembly. The wrong default made RobotDoc() return Nothing, and because Apply
    ' is allowed to find no robot assembly at all (MainForm.RobotDoc, complain:=False),
    ' ToolSwitcher was skipped in silence instead of reporting anything.
    Public Function RobotAssemblyPath() As String
        Dim p As String = Setting("RobotAssemblyPath", "")
        If String.IsNullOrEmpty(p) Then p = CadFile(CadRoot(), "05_robot", "robot_mainAssembly.CATProduct")
        Return p
    End Function

    Public Function OpenTableAssembly() As Boolean
        Return String.Equals(Setting("OpenTableAssembly", "true"), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    ' main_assembly = table_mainAssembly + robot_mainAssembly. When open, Apply updates
    ' and brings THIS to the front, so the robot visibly repositions/rotates - the table
    ' assembly alone does not contain the robot. The table switchers still run on the
    ' table document (loaded as a component), then main_assembly is updated on top.
    Public Function MainAssemblyPath() As String
        Dim p As String = Setting("MainAssemblyPath", "")
        If String.IsNullOrEmpty(p) Then p = CadFile(CadRoot(), "00_main", "main_assembly.CATProduct")
        Return p
    End Function

    Public Function OpenMainAssembly() As Boolean
        Return String.Equals(Setting("OpenMainAssembly", "true"), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    ' Nominal arm segments, read from the RoboDK model of the Fanuc LR Mate 200iC.
    ' The task allows nominal .. nominal + 20 %, and the UI derives the slider range from
    ' these, so a different robot only needs a change in the .config file.
    Public Function NominalLenJ2J3() As Double
        Return Number("NominalLenJ2J3", 300.0)
    End Function

    Public Function NominalLenJ3J5() As Double
        Return Number("NominalLenJ3J5", 320.0)
    End Function

    ' How far the robot base body reaches past the outer mounting bolts, per side.
    ' Default is the Fanuc LR Mate 200iC: roughly a 190 x 150 mm base on the 100 x 60 mm
    ' bolt raster. It is an app-side limit only - no such parameter exists in the master.
    Public Function RobotBaseOverhang() As Double
        Return Number("RobotBaseOverhang", 45.0)
    End Function

    ' NO RobotModel KEY (decision 17.08.2026): only the assigned FANUC LR Mate 200iC is
    ' modelled, so there is nothing to switch between. Its data lives in RobotCatalog.

    ' The master's hole sketch used to apply axisOffsetX/Y rotated by 90 degrees; it was
    ' corrected on 15.08.2026, so this defaults to true. Setting it back to false makes
    ' the app keep both base offsets valid for BOTH axes - the safety net to reach for if
    ' the raster ever lands sideways again.
    Public Function RobotMountSketchFixed() As Boolean
        Return String.Equals(Setting("RobotMountSketchFixed", "true"), "true", StringComparison.OrdinalIgnoreCase)
    End Function

    Private Function Number(ByVal key As String, ByVal dflt As Double) As Double
        Dim d As Double
        If Double.TryParse(Setting(key, ""), Globalization.NumberStyles.Float,
                           Globalization.CultureInfo.InvariantCulture, d) Then
            If d > 0 Then Return d
        End If
        Return dflt
    End Function

End Module
