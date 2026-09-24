Option Explicit On
Option Strict Off

Imports System.Runtime.InteropServices
Imports System.Text

' Re-imports the robot's DMU mechanisms into main_assembly at the end of Apply.
'
' WHY THIS HAS TO EXIST
'
' main_assembly does not own a mechanism of its own; it carries IMPORTED copies of the
' ones in robot_mainAssembly ("Robot") and in the gripper. An import holds on to the joint
' geometry it saw when it was made, so as soon as Apply changes anything the kinematic
' points depend on (segment lengths, robot mount, tool), the imported "Robot" drops to
' DOF=Undefined. robot_mainAssembly itself stays fine. The fix by hand was always the
' same: delete the imported mechanisms, then DMU Kinematics > Import Sub-Mechanisms.
'
' MechanismRepair.CleanupImportedMechanisms already did the first half before the update.
' This is the second half, after the update, when the geometry is final.
'
' HOW  (measured 24.09.2026)
'
' The Automation API has no import call - not in KinTypeLib, not on Mechanisms. What
' works is the command itself: in the DMUKinematics workbench, with main_assembly's root
' selected, StartCommand "Import Sub-Mechanisms" imports every sub-mechanism and then
' shows a modal "Sub-mechanisms Import Result" box ("2 Sub-mechanisms have been imported
' successfully."). StartCommand returns at once, so the box is waited for and its OK
' pressed from here. The command name is the English title - another CATIA language
' would need its own.
Public Module MechanismImport

    Private Const COMMAND_NAME As String = "Import Sub-Mechanisms"
    Private Const RESULT_TITLE As String = "Sub-mechanisms Import Result"
    Private Const WAIT_MS As Integer = 30000

    Private Delegate Function EnumProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    <DllImport("user32.dll")>
    Private Function EnumWindows(ByVal cb As EnumProc, ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function EnumChildWindows(ByVal parent As IntPtr, ByVal cb As EnumProc, ByVal lParam As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function IsWindowVisible(ByVal hWnd As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll")>
    Private Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef pid As UInteger) As UInteger
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Function SendMessage(ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As StringBuilder) As IntPtr
    End Function

    <DllImport("user32.dll")>
    Private Function PostMessage(ByVal hWnd As IntPtr, ByVal msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    End Function

    Private Const WM_GETTEXT As UInteger = &HD
    Private Const BM_CLICK As UInteger = &HF5


    ' Deletes whatever imported mechanisms main_assembly still has, imports them again
    ' and reports the DOF of each. Returns a line for the Apply report.
    '
    ' NEVER THROWS. The parameters are written and the model rebuilt by now; a mechanism
    ' that did not come back is something to tell the user, not a failed Apply.
    Public Function Reimport(ByVal app As Object, ByVal mainDoc As Object) As String
        If app Is Nothing OrElse mainDoc Is Nothing Then Return ""

        Dim root As Object = Nothing
        Try
            root = mainDoc.Product
        Catch
            Return ""
        End Try

        Dim wbBefore As String = ""
        Try
            CatiaConnector.Activate(mainDoc)
            wbBefore = CStr(app.GetWorkbenchId())
        Catch
        End Try

        Try
            ' Leftovers first: a second import beside a stale one gives two "Robot"s.
            DeleteImported(mainDoc, root)

            If wbBefore <> "DMUKinematics" Then app.StartWorkbench("DMUKinematics")

            Dim sel As Object = mainDoc.Selection
            sel.Clear()
            sel.Add(root)
            app.StartCommand(COMMAND_NAME)

            Dim said As String = WaitAndConfirm()

            Dim result As String = Summary(root)
            If said = "" Then
                result = "DMU: no confirmation from '" & COMMAND_NAME & "' within " &
                         (WAIT_MS \ 1000) & " s. Mechanisms now: " & result &
                         ". If Robot is missing, run DMU Kinematics > Import Sub-Mechanisms by hand."
            Else
                result = "DMU: " & said & "  " & result
            End If
            Return result

        Catch ex As Exception
            Return "DMU: re-import failed (" & ex.Message & "). In main_assembly: delete the " &
                   "imported mechanisms, then DMU Kinematics > Import Sub-Mechanisms."
        Finally
            Try
                If wbBefore <> "" AndAlso wbBefore <> "DMUKinematics" Then app.StartWorkbench(wbBefore)
            Catch
            End Try
        End Try
    End Function


    Private Sub DeleteImported(ByVal mainDoc As Object, ByVal root As Object)
        Dim mechs As Object = root.GetTechnologicalObject("Mechanisms")
        Dim n As Integer = CInt(mechs.Count)
        If n = 0 Then Return
        Dim sel As Object = mainDoc.Selection
        sel.Clear()
        For i As Integer = n To 1 Step -1
            sel.Add(mechs.Item(i))
        Next
        sel.Delete()
        sel.Clear()
    End Sub


    ' "Robot DOF=6, Mechanism.1 DOF=1" - the short name of each mechanism with its DOF,
    ' read back rather than assumed.
    Private Function Summary(ByVal root As Object) As String
        Dim parts As New List(Of String)
        Try
            Dim mechs As Object = root.GetTechnologicalObject("Mechanisms")
            For i As Integer = 1 To CInt(mechs.Count)
                Dim m As Object = mechs.Item(i)
                Dim nm As String = CStr(m.Name)
                If nm.Contains("\") Then nm = nm.Substring(nm.LastIndexOf("\"c) + 1)
                Dim dof As String
                Try
                    dof = CStr(CInt(m.NbDof))
                Catch
                    dof = "Undefined"
                End Try
                parts.Add(nm & " DOF=" & dof)
            Next
        Catch
        End Try
        If parts.Count = 0 Then Return "(none)"
        Return String.Join(", ", parts)
    End Function


    ' Makes robot_mainAssembly.1 a RIGID sub-assembly of main_assembly again, if it is not.
    '
    ' WHY  (measured 24.09.2026)
    '
    ' The robot is placed by one constraint, BASE_ORIGIN (rb_base) on robot_mount_origin
    ' (master). With the robot FLEXIBLE in main_assembly, a change of table height made the
    ' solver move rb_base ALONE: the table went from 1200 to 1416, rb_base followed inside
    ' the robot, the arm and robot_mainAssembly.1 stayed at 1200 - and the constraint still
    ' reported itself satisfied. Rigid, the same update lifts the whole robot. Importing the
    ' sub-mechanisms does NOT make it flexible again, so once rigid it normally stays so;
    ' this is the guard for when it is switched back by hand.
    '
    ' The API cannot read or set flexibility. A flexible sub-assembly is recognised by its
    ' own constraints showing up in main_assembly as copies ("Coincidence.60.1": both
    ' elements inside robot_mainAssembly.1), and switched with the command
    ' "Flexible/Rigid Sub-Assembly", whose "Rigification warning" is confirmed here.
    Public Function EnsureRobotRigid(ByVal app As Object, ByVal mainDoc As Object,
                                     ByVal robotInstance As String) As String
        If app Is Nothing OrElse mainDoc Is Nothing Then Return ""
        Try
            Dim root As Object = mainDoc.Product
            If Not IsFlexible(root, robotInstance) Then Return ""

            CatiaConnector.Activate(mainDoc)
            Dim wbBefore As String = CStr(app.GetWorkbenchId())
            If wbBefore <> "Assembly" Then app.StartWorkbench("Assembly")

            Dim sel As Object = mainDoc.Selection
            sel.Clear()
            sel.Add(root.Products.Item(robotInstance))
            app.StartCommand("Flexible/Rigid Sub-Assembly")
            WaitAndConfirm("Rigification warning")
            sel.Clear()

            If wbBefore <> "" AndAlso wbBefore <> "Assembly" Then app.StartWorkbench(wbBefore)

            If IsFlexible(root, robotInstance) Then
                Return "Robot: still a FLEXIBLE sub-assembly of main_assembly, so it will not " &
                       "follow the table. Right-click " & robotInstance & " > Components > " &
                       "Flexible/Rigid Sub-Assembly."
            End If
            Return "Robot: made a rigid sub-assembly of main_assembly again (it was flexible, " &
                   "so only rb_base followed the table)."
        Catch ex As Exception
            Return "Robot: could not check flexible/rigid (" & ex.Message & ")."
        End Try
    End Function


    Private Function IsFlexible(ByVal root As Object, ByVal robotInstance As String) As Boolean
        Dim cons As Object = AssemblySwap.GetConstraints(root)
        If cons Is Nothing Then Return False
        Dim inside As String = "/" & robotInstance & "/"
        For i As Integer = 1 To CInt(cons.Count)
            Dim e1 As String = AssemblySwap.ElementName(cons.Item(i), 1)
            Dim e2 As String = AssemblySwap.ElementName(cons.Item(i), 2)
            If e1.Contains(inside) AndAlso e2.Contains(inside) Then Return True
        Next
        Return False
    End Function


    ' Waits for the result box, returns its message and presses OK. "" on timeout.
    Private Function WaitAndConfirm(Optional ByVal title As String = RESULT_TITLE) As String
        Dim procs() As Diagnostics.Process = Diagnostics.Process.GetProcessesByName("CNEXT")
        If procs.Length = 0 Then Return ""
        Dim pid As UInteger = CUInt(procs(0).Id)

        Dim sw As Diagnostics.Stopwatch = Diagnostics.Stopwatch.StartNew()
        While sw.ElapsedMilliseconds < WAIT_MS
            Dim box As IntPtr = FindTopWindow(pid, title)
            If box <> IntPtr.Zero Then
                Dim message As String = ""
                Dim okButton As IntPtr = IntPtr.Zero
                EnumChildWindows(box,
                    Function(h As IntPtr, l As IntPtr) As Boolean
                        Dim t As String = TextOf(h)
                        If t = "OK" Then
                            okButton = h
                        ElseIf t.Length > message.Length Then
                            message = t
                        End If
                        Return True
                    End Function, IntPtr.Zero)
                If okButton <> IntPtr.Zero Then PostMessage(okButton, BM_CLICK, IntPtr.Zero, IntPtr.Zero)
                Return If(message = "", "imported.", message.Trim())
            End If
            Threading.Thread.Sleep(250)
        End While
        Return ""
    End Function


    Private Function FindTopWindow(ByVal pid As UInteger, ByVal title As String) As IntPtr
        Dim found As IntPtr = IntPtr.Zero
        EnumWindows(
            Function(h As IntPtr, l As IntPtr) As Boolean
                Dim p As UInteger = 0
                GetWindowThreadProcessId(h, p)
                If p = pid AndAlso IsWindowVisible(h) AndAlso TextOf(h) = title Then
                    found = h
                    Return False
                End If
                Return True
            End Function, IntPtr.Zero)
        Return found
    End Function


    Private Function TextOf(ByVal h As IntPtr) As String
        Dim sb As New StringBuilder(1024)
        SendMessage(h, WM_GETTEXT, New IntPtr(1024), sb)
        Return sb.ToString()
    End Function

End Module
