Option Explicit On
Option Strict Off

Imports System.Text

' Swaps the whole tool when the user changes the object type.
'
' Task section 9 asks that choosing object 1 or object 2 reconfigure the adapter plate,
' the gripper, the fingers and the object itself. Knowledge cannot do that: a rule can
' set a number, but replacing a component is a change of geometry. So R_Tool states what
' should be fitted and this fits it - the same division of labour as R_Profile with
' ProfileSwitcher and gusset size with GussetSwitcher, and the same one the report should
' describe as "what the application does that the model cannot".
'
' ONE SWAP, NOT FOUR
' The four things that change are packed into a sub-assembly, sub_tool_1 or sub_tool_2,
' which expose an identical set of publications. So this module replaces a single
' component and rebuilds a single constraint. Four separate swaps would mean four
' constraints to rebuild and four ways to end up half-swapped; this way the failure mode
' is "nothing happened", which is recoverable.
'
' THE PRECONDITION THAT BIT THE GUSSETS, AND WILL BITE THIS
' The tool instance must carry its OWN constraint, addressed through a publication. An
' instance with no constraint of its own - a Reuse Pattern copy, say - cannot be
' reconnected after the component is exchanged, and would come back unconstrained. There
' is only one tool, so this is easy to satisfy: constrain it, once, by hand.
'
' EXACTLY ONE CONSTRAINT, WHICH IS WHY THE PUBLICATION IS AN AXIS SYSTEM
' AssemblySwap reads ONE constraint (ReadConstraintOf returns the first that mentions the
' instance) and rebuilds ONE. Mating on an axis line plus a face plane would need two -
' and would still leave the tool free to spin about its own axis, because a line
' coincidence plus a plane coincidence is 5 constrained DOF, not 6. So both sub_tool_x
' publish a single AXIS SYSTEM, TOOL_AXS, mated to rb_link6/!kin/FLANGE_AXS: one
' constraint, zero DOF, no code here that has to loop.
'
' Recognition is by INSTANCE NAME, not by part number, for the same reason
' GussetSwitcher recognises corner brackets by name: the name is ours to control and
' survives the swap (SwapOne restores it), whereas the part number is what we are
' changing.

Public Module ToolSwitcher

    ' The instance in robot_mainAssembly. Name it exactly this when you insert the tool.
    Public Const INSTANCE_NAME As String = "tool_1"

    ' The root publication both sub_tool_x expose, and the way back if the resolved
    ' reference does not survive the swap. Both documents must spell it identically.
    Public Const PUBLICATION As String = "TOOL_AXS"

    ' Part numbers of the two sub-assemblies.
    Private Const PN_PREFIX As String = "sub_tool_"

    ' Where the documents live, relative to CadRoot - which ALREADY ends in 01_cad
    ' (AppConfig.CadRoot). This used to read "01_cad\06_tool" and resolved to
    ' ...\01_cad\01_cad\06_tool, so the file was never found and the swap reported
    ' "does not exist". Every other switcher spells the folder relative to CadRoot
    ' (ScrewSwitcher "04_fasteners", GussetSwitcher "01_table"); this is the same.
    Private Const FOLDER As String = "06_tool"


    Public Function TargetPartNumber(ByVal objectType As String) As String
        Select Case Trim(If(objectType, ""))
            Case "1" : Return PN_PREFIX & "1"
            Case "2" : Return PN_PREFIX & "2"
        End Select
        Return ""
    End Function


    ' Which sub_tool is ACTUALLY fitted right now: "1", "2", or "" when the tool instance
    ' cannot be found. Reads, changes nothing.
    '
    ' The model can hold two different answers to "which gripper is this". object_type in
    ' the master is what R_Tool branches on and what the window shows; the part number of
    ' the tool_1 instance is what is physically in the assembly. They agree only for as
    ' long as nobody swaps one without the other - and on 29.08.2026 they did not: the
    ' master said "1" while sub_tool_2 was mounted, so the window read back the wrong
    ' gripper and there was nothing on screen to say why.
    Public Function FittedObjectType(ByVal robotDoc As Object) As String
        If robotDoc Is Nothing Then Return ""

        Dim root As Object = Nothing
        Try
            root = robotDoc.Product
        Catch
            Return ""
        End Try
        If root Is Nothing Then Return ""

        Dim kids As Object
        Dim count As Integer = 0
        Try
            kids = root.Products
            count = CInt(kids.Count)
        Catch
            Return ""
        End Try

        For i As Integer = 1 To count
            Dim pn As String = ""
            Dim nm As String = ""
            Try
                pn = CStr(kids.Item(i).PartNumber)
                nm = CStr(kids.Item(i).Name)
            Catch
                Continue For
            End Try

            If nm.StartsWith(INSTANCE_NAME, StringComparison.OrdinalIgnoreCase) OrElse
               pn.StartsWith(PN_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                If pn.EndsWith("1") Then Return "1"
                If pn.EndsWith("2") Then Return "2"
                Return ""
            End If
        Next
        Return ""
    End Function


    Public Function TargetPath(ByVal cadRoot As String, ByVal objectType As String) As String
        Dim pn As String = TargetPartNumber(objectType)
        If pn = "" Then Return ""
        Return AppConfig.CadFile(cadRoot, FOLDER, pn & ".CATProduct")
    End Function


    ' Main entry. Returns the message for the user; throws only when the assembly is
    ' unreachable.
    '
    ' The three phases are the safety contract, copied from GussetSwitcher because it is
    ' the one that survived contact with a real assembly:
    '   1) read and understand everything, touching nothing;
    '   2) open the replacement document;
    '   3) swap, and stop at the first failure.
    Public Function EnsureTool(ByVal app As Object, ByVal robotDoc As Object,
                               ByVal objectType As String,
                               ByVal cadRoot As String) As String

        Dim want As String = TargetPartNumber(objectType)
        If want = "" Then
            Return "Tool: nothing to swap - object_type = " & Quote(objectType) &
                   " is neither ""1"" nor ""2""."
        End If

        Dim root As Object = Nothing
        Try
            root = robotDoc.Product
        Catch
            root = Nothing
        End Try
        If root Is Nothing Then Throw New Exception("The robot assembly has no Product.")

        ' --- 1) find the tool instance -----------------------------------------------------
        Dim kids As Object = root.Products
        Dim count As Integer = 0
        Try
            count = CInt(kids.Count)
        Catch
            count = 0
        End Try

        Dim instName As String = ""
        Dim havePn As String = ""
        Dim seen As New StringBuilder()

        For i As Integer = 1 To count
            Dim pn As String = ""
            Dim nm As String = ""
            Try
                pn = CStr(kids.Item(i).PartNumber)
                nm = CStr(kids.Item(i).Name)
            Catch
                Continue For
            End Try
            seen.AppendLine("    " & nm & "   (" & pn & ")")

            If nm.StartsWith(INSTANCE_NAME, StringComparison.OrdinalIgnoreCase) OrElse
               pn.StartsWith(PN_PREFIX, StringComparison.OrdinalIgnoreCase) Then
                instName = nm
                havePn = pn
                Exit For
            End If
        Next

        If instName = "" Then
            Return "Tool: no instance named " & INSTANCE_NAME & " (and no " & PN_PREFIX &
                   "* part number) in the robot assembly." & vbCrLf & vbCrLf &
                   "What is in there:" & vbCrLf & seen.ToString().TrimEnd() & vbCrLf & vbCrLf &
                   "Insert sub_tool_1.CATProduct on the flange and rename the instance " &
                   "to " & INSTANCE_NAME & " - TODO.md B5."
        End If

        If String.Equals(havePn, want, StringComparison.OrdinalIgnoreCase) Then
            Return "Tool: already " & want & " (" & GripperCatalog.ByType(objectType).Caption & ")."
        End If

        ' --- 2) read the constraint before touching anything -------------------------------
        Dim f As AssemblySwap.SwapTarget = AssemblySwap.ReadConstraintOf(root, instName)
        If f Is Nothing Then
            Return "Tool NOT swapped - the assembly was left untouched." & vbCrLf &
                   "The constraint of " & instName & " could not be read." & vbCrLf & vbCrLf &
                   AssemblySwap.DumpConstraints(root) & vbCrLf &
                   "Swap it by hand: right-click " & instName & " > Components >" & vbCrLf &
                   "Replace Component > " & want & ".CATProduct."
        End If
        f.InstanceName = instName
        f.PartNumber = havePn
        f.WantPartNumber = want
        f.PubName = PUBLICATION

        ' --- 3) open the replacement --------------------------------------------------------
        Dim path As String = TargetPath(cadRoot, objectType)
        If Not IO.File.Exists(path) Then
            Return "Tool NOT swapped: " & path & " does not exist." & vbCrLf &
                   "Build the two tool sub-assemblies first - TODO.md B4."
        End If

        Dim srcDoc As Object = CatiaConnector.GetOrOpenDocument(app, path)
        Dim srcProd As Object = Nothing
        Try
            srcProd = srcDoc.Product
        Catch
            srcProd = Nothing
        End Try
        If srcProd Is Nothing Then
            Return "Tool NOT swapped: " & path & " opened but has no Product."
        End If

        ' --- 4) swap ------------------------------------------------------------------------
        Try
            AssemblySwap.SwapOne(root, f, srcProd, TranslatorFor(f))
        Catch ex As Exception
            Return "Tool NOT swapped: " & ex.Message & vbCrLf & vbCrLf &
                   "The component may already be exchanged with its constraint missing. " &
                   "Check " & instName & " in the tree; if it is loose, constrain it by " &
                   "hand: one Coincidence, " & INSTANCE_NAME & "/TOOL_AXS on " &
                   "rb_link6/!kin/FLANGE_AXS."
        End Try

        ' --- 5) put the tool back into the mechanism ---------------------------------------
        ' The swap removed the tool_1 Product and added a new one, so the DMU Rigid joint
        ' that held the tool to the flange now points at an object that no longer exists.
        ' Nothing in AssemblySwap knows the mechanism is there - it cannot, it is used for
        ' gussets and screws too - so the repair belongs here, next to the one swap that
        ' has a joint on it.
        '
        ' Left undone, this is invisible until somebody runs a simulation and CATIA
        ' refuses with "some joints are invalid". HOWTO_dmu_kinematics.md 5.3 raised the
        ' question when the joint was first made and it stayed open until 12e measured
        ' that a joint can be removed and a Rigid one created through the API.
        Dim repair As String = MechanismRepair.RepairToolJoint(robotDoc, root, instName)

        Dim g As GripperInfo = GripperCatalog.ByType(objectType)
        Dim msg As New StringBuilder()
        msg.AppendLine("Tool: " & havePn & " -> " & want & "  (" & g.Caption & ")")
        msg.AppendLine("  object " & g.ObjA & " x " & g.ObjB & " x " & g.ObjC &
                       " mm, stroke " & g.StrokePerJaw & " mm per jaw")
        If f.UsedRef <> "" Then msg.AppendLine("  via " & f.UsedRef)
        If repair <> "" Then msg.AppendLine(repair)
        If Not g.Verified Then
            msg.AppendLine("  note: this gripper's dimensions are not verified against " &
                           "the CAD yet - TODO2.md M3, M4.")
        End If
        Return msg.ToString().TrimEnd()
    End Function


    ' A path rewriter for the swap.
    '
    ' DisplayName gives the RESOLVED path, so a reference to the tool's own side reads
    ' something like
    '     robot/tool_1.1/sub_tool_1.1/!TOOL_AXS
    ' and the document name inside it has to be translated, exactly as GussetSwitcher
    ' translates GUSSET_45X45 -> GUSSET_60X60. The INSTANCE name is not translated: it is
    ' restored after the swap, so it stays literally what it was.
    '
    ' Publications keep the same names in both sub-assemblies on purpose - that is what
    ' makes the two tools interchangeable - so only the part number token changes.
    Private Function TranslatorFor(ByVal f As AssemblySwap.SwapTarget) As Func(Of String, String)
        Dim a As String = f.PartNumber
        Dim b As String = f.WantPartNumber
        If String.IsNullOrEmpty(a) OrElse String.IsNullOrEmpty(b) OrElse a = b Then
            Return Function(p As String) p
        End If
        Return Function(p As String) p.Replace(a, b)
    End Function


    Private Function Quote(ByVal s As String) As String
        If s Is Nothing Then Return """"""
        Return """" & s & """"
    End Function

End Module
