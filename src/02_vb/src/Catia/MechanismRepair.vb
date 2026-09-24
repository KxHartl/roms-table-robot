Option Explicit On
Option Strict Off

Imports System.Text

' Puts the tool back into the DMU mechanism after the tool has been swapped.
'
' WHY THIS HAS TO EXIST
'
' AssemblySwap.SwapOne exchanges a component the only way the V5 API allows: Products has
' no Replace and ReferenceProduct is read-only, so it removes the instance and adds a new
' one. That is fine for the assembly - the constraint is rebuilt and read back - but the
' DMU mechanism holds its own reference to the Product object that was removed, and
' nothing in the swap knows the mechanism exists.
'
' The result is a joint pointing at nothing. It looks perfectly ordinary in the tree.
' HOWTO_dmu_kinematics.md section 5.3 records what that costs: the mechanism reports
' DOF=Undefined and "Simulation with Laws" refuses with "some joints are invalid (missing
' part or geometry)" - and only at simulation time, long after the swap that caused it.
'
' Adding a fresh Rigid joint beside the broken one is not a repair. The broken one stays,
' and so does the refusal. It has to be removed first.
'
' WHAT WAS MEASURED BEFORE ANY OF THIS WAS WRITTEN  (12e_probe_nested_links, 29.08.2026)
'
'   Joints.Remove(i)                        -> works; joint count really drops
'   AddJoint("CATKinRigidJoint", 2 refs)    -> works, with references to two PUBLISHED
'                                              axis systems. Two references to the
'                                              components themselves are refused.
'   reference spelling                      -> "<root>/<instance>/!<PUBLICATION>"
'                                              The bare geometry path (!kin/...) is not
'                                              needed; publications resolve.
'
' 12_kin_mechanism.CATScript's header says a MECHANISM cannot be deleted safely through
' the API. That turned out to be a statement about mechanisms only - a single joint comes
' out cleanly, which is what makes this module possible at all.
Public Module MechanismRepair

    ' The name this module gives the joint it creates, so that the NEXT repair can find it
    ' without guessing. The first joint was made by hand in the GUI and got CATIA's
    ' automatic name, so that one is recognised too.
    Public Const TOOL_JOINT_NAME As String = "TOOL_RIGID"

    ' Tried in order. Recognises TOOL_RIGID (standard), Rigid.7 (user's model),
    ' Rigid.8 (baseline model), and standard CATIA names.
    Private ReadOnly CANDIDATE_NAMES As String() = {TOOL_JOINT_NAME, "Rigid.7", "Rigid.8", "Rigid.1", "Rigid.2", "Rigid.9"}

    ' The two published axis systems the joint is built over. They identify the two parts;
    ' the relative position it freezes is whatever the assembly constraint has already
    ' put them in, which is why the update below is not optional.
    Private Const FLANGE_INSTANCE As String = "rb_link6.1"
    Private Const FLANGE_PUB As String = "FLANGE_AXS"
    Private Const TOOL_PUB As String = "TOOL_AXS"


    ' Repairs the tool joint after a swap. Returns a line for the caller's report, or ""
    ' when there is nothing to do.
    '
    ' NEVER THROWS. A tool swap that succeeded must not be reported as a failure because
    ' the mechanism could not be touched - the assembly is correct either way, and what
    ' the user needs is to be told precisely what state the mechanism is in.
    Public Function RepairToolJoint(ByVal robotDoc As Object, ByVal root As Object,
                                    ByVal toolInstance As String) As String

        Dim mech As Object = FindMechanism(robotDoc)
        If mech Is Nothing Then
            ' No mechanism yet. That is the ordinary case while the model is still being
            ' built, and it is not a problem: the joint will be made against whichever
            ' tool is mounted when the mechanism is created.
            Return ""
        End If

        Dim before As Integer = SafeInt(mech, "NbJoints")

        ' --- find the joint that used to hold the tool ------------------------------------
        Dim idx As Integer = FindJoint(mech, CANDIDATE_NAMES)
        If idx <= 0 Then
            ' Deliberately does NOT create one. A second Rigid beside an unrecognised
            ' first would leave two joints on one pair of parts, which is a worse model
            ' than the one we started with and harder to diagnose.
            Return "  ! the mechanism's tool joint was not found by name (" &
                   String.Join(", ", CANDIDATE_NAMES) & ")." & vbCrLf &
                   "    Joints present: " & JointNames(mech) & vbCrLf &
                   "    The tool is swapped but the mechanism still points at the old " &
                   "component. Rebuild the joint by hand: Insert > New Joint > Rigid, " &
                   "rb_link6.1 and " & toolInstance & ", and check BOTH fields show a " &
                   "name before OK."
        End If

        Dim oldName As String = SafeName(mech.Joints.Item(idx))

        ' --- remove it --------------------------------------------------------------------
        Dim why As String = ""
        Try
            mech.Joints.Remove(idx)
        Catch ex As Exception
            why = ex.Message
        End Try

        Dim afterRemove As Integer = SafeInt(mech, "NbJoints")
        If afterRemove >= before Then
            Return "  ! could not remove the old tool joint " & oldName &
                   If(why = "", "", " (" & why & ")") & "." & vbCrLf &
                   "    The mechanism still points at the component that was replaced, " &
                   "so a simulation will refuse with ""some joints are invalid""." &
                   vbCrLf & "    Delete " & oldName & " in the tree and remake it: " &
                   "Insert > New Joint > Rigid, rb_link6.1 and " & toolInstance & "."
        End If

        ' --- update BEFORE rebuilding -----------------------------------------------------
        ' A Rigid joint freezes the relative position the two parts are in when it is
        ' made. The swap has already rebuilt the Coincidence, but until the assembly is
        ' updated the new tool still sits where AddComponent dropped it - so making the
        ' joint first would weld the tool into the wrong place, and the arm would carry it
        ' there for the whole simulation.
        Try
            root.Update()
        Catch
            ' Reported through the verification below rather than aborting: an update that
            ' fails here is usually a constraint problem the swap already described.
        End Try

        ' --- rebuild ----------------------------------------------------------------------
        Dim r1 As Object = MakeRef(root, FLANGE_INSTANCE, FLANGE_PUB)
        Dim r2 As Object = MakeRef(root, toolInstance, TOOL_PUB)
        If r1 Is Nothing OrElse r2 Is Nothing Then
            Return "  ! removed " & oldName & " but could not build the references for " &
                   "the new joint." & vbCrLf &
                   "    Remake it by hand: Insert > New Joint > Rigid, " &
                   FLANGE_INSTANCE & " and " & toolInstance & "."
        End If

        Dim refs(1) As Object
        refs(0) = r1
        refs(1) = r2

        Dim joint As Object = Nothing
        why = ""
        Try
            joint = mech.AddJoint("CATKinRigidJoint", refs)
        Catch ex As Exception
            why = ex.Message
        End Try

        If joint Is Nothing Then
            Return "  ! removed " & oldName & " but AddJoint(CATKinRigidJoint) failed" &
                   If(why = "", "", ": " & why) & "." & vbCrLf &
                   "    The mechanism is now one joint short - the tool would hang in " &
                   "mid-air during a simulation." & vbCrLf &
                   "    Remake it by hand: Insert > New Joint > Rigid, " &
                   FLANGE_INSTANCE & " and " & toolInstance & ", and check BOTH fields " &
                   "show a name before OK."
        End If

        Try
            joint.Name = TOOL_JOINT_NAME
        Catch
            ' A joint that will not be renamed still works; the next repair falls back to
            ' whatever CATIA called it, which is why the candidate list exists.
        End Try

        ' --- verify, rather than assume ----------------------------------------------------
        Dim after As Integer = SafeInt(mech, "NbJoints")
        Dim dof As String = SafeDof(mech)

        Dim msg As New StringBuilder()
        msg.AppendLine("  mechanism: " & oldName & " rebuilt as " & SafeName(joint) &
                       "  (joints " & before & " -> " & after & ", DOF " & dof & ")")

        If after <> before Then
            msg.AppendLine("    ! the joint count did not come back to " & before &
                           " - check the mechanism in the tree before simulating.")
        ElseIf dof <> "0" Then
            ' Measured 29.08.2026: NbDof does not always refresh straight after a change,
            ' so this is a flag to check, not a verdict. Claiming the model is broken on
            ' a stale counter would be as wrong as claiming it is fine.
            msg.AppendLine("    note: DOF reads " & dof & " rather than 0. It may simply " &
                           "not have refreshed; confirm with one Simulation with Laws " &
                           "before recording.")
        End If

        Return msg.ToString().TrimEnd()
    End Function


    ' The mechanism of a product document, or Nothing. Never throws: this is called to
    ' decide whether there is anything to repair, and an exception there would turn a tool
    ' swap that worked into a dialog about DMU.
    '
    ' The entry point is GetTechnologicalObject("Mechanisms"), NOT GetWorkbench - proved on
    ' this installation by 09b_probe_kinematics, and it is what Dassault's own sample
    ' CAAKiiMechanismSolve.CATScript inside V5Automation.chm uses.
    '
    ' It lived in MechanismDriver until 30.08.2026; when the Simulate button went, this was
    ' the one part of that module still in use.
    Private Function FindMechanism(ByVal prodDoc As Object) As Object
        If prodDoc Is Nothing Then Return Nothing

        Dim root As Object = Nothing
        Try
            root = prodDoc.Product
        Catch
            Return Nothing
        End Try
        If root Is Nothing Then Return Nothing

        Dim mechs As Object = Nothing
        Try
            mechs = root.GetTechnologicalObject("Mechanisms")
        Catch
            Return Nothing
        End Try
        If mechs Is Nothing Then Return Nothing

        Try
            If CInt(mechs.Count) < 1 Then Return Nothing
            Return mechs.Item(1)
        Catch
            Return Nothing
        End Try
    End Function


    ' Index of the first joint whose name matches one of `names`, or 0.
    Private Function FindJoint(ByVal mech As Object, ByVal names As String()) As Integer
        Dim n As Integer = SafeInt(mech, "NbJoints")
        If n < 1 Then Return 0

        For Each want As String In names
            For i As Integer = 1 To n
                Dim nm As String = ""
                Try
                    nm = CStr(mech.Joints.Item(i).Name)
                Catch
                    Continue For
                End Try
                If String.Equals(nm, want, StringComparison.OrdinalIgnoreCase) Then Return i
            Next
        Next

        ' Fallback: find any joint whose name contains "TOOL" or starts with "Rigid"
        For i As Integer = 1 To n
            Dim nm As String = ""
            Try
                nm = CStr(mech.Joints.Item(i).Name)
            Catch
                Continue For
            End Try
            If nm.IndexOf("TOOL", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
               nm.StartsWith("Rigid", StringComparison.OrdinalIgnoreCase) Then
                Return i
            End If
        Next

        Return 0
    End Function


    ' Every joint's name, so a failure report names what IS there instead of only what is
    ' missing - the difference between a message that ends the problem and one that starts
    ' another round of looking.
    Private Function JointNames(ByVal mech As Object) As String
        Dim n As Integer = SafeInt(mech, "NbJoints")
        Dim parts As New List(Of String)
        For i As Integer = 1 To n
            Try
                parts.Add(CStr(mech.Joints.Item(i).Name))
            Catch
                parts.Add("?")
            End Try
        Next
        If parts.Count = 0 Then Return "(none)"
        Return String.Join(", ", parts)
    End Function


    ' "<root>/<instance>/!<publication>" - the one spelling 12e proved AddJoint accepts.
    ' CreateReferenceFromName never fails, so a returned object is not evidence; the
    ' verdict comes from AddJoint, which is why this is not checked here.
    Private Function MakeRef(ByVal root As Object, ByVal instance As String,
                             ByVal pub As String) As Object
        Dim rootName As String = SafeName(root)
        Try
            Return root.CreateReferenceFromName(rootName & "/" & instance & "/!" & pub)
        Catch
            Return Nothing
        End Try
    End Function


    ' DOF as a STRING, so "cannot be read" is a value rather than an exception - which is
    ' precisely the state a joint with a dangling link puts the mechanism in.
    Private Function SafeDof(ByVal mech As Object) As String
        Try
            Return CStr(CInt(mech.NbDof))
        Catch
            Return "?"
        End Try
    End Function


    Private Function SafeInt(ByVal o As Object, ByVal prop As String) As Integer
        Try
            If prop = "NbJoints" Then Return CInt(o.NbJoints)
            If prop = "NbCommands" Then Return CInt(o.NbCommands)
        Catch
        End Try
        Return 0
    End Function


    ''' <summary>
    ''' Cleans up imported DMU mechanisms and conflicting Fix constraints from main_assembly.
    ''' When a mechanism is simulated in main_assembly, CATIA imports/projects the mechanism
    ''' and automatically creates a Fix constraint on rb_base.1. After tool swaps or edits,
    ''' these imported mechanisms hold stale links and the Fix constraint fights assembly mounting.
    ''' Removing them leaves main_assembly clean so CATIA can cleanly re-import upon simulation.
    ''' </summary>
    Public Function CleanupImportedMechanisms(ByVal mainDoc As Object) As String
        If mainDoc Is Nothing Then Return ""

        Dim root As Object = Nothing
        Try
            root = mainDoc.Product
        Catch
            Return ""
        End Try
        If root Is Nothing Then Return ""

        Dim msg As New StringBuilder()

        ' 1) Remove mechanisms in main_assembly via Selection.Delete
        Try
            Dim mechs As Object = root.GetTechnologicalObject("Mechanisms")
            If mechs IsNot Nothing Then
                Dim nMechs As Integer = 0
                Try
                    nMechs = CInt(mechs.Count)
                Catch
                    nMechs = 0
                End Try

                If nMechs > 0 Then
                    Dim sel As Object = Nothing
                    Try
                        sel = mainDoc.Selection
                    Catch
                        sel = Nothing
                    End Try

                    If sel IsNot Nothing Then
                        sel.Clear()
                        For i As Integer = nMechs To 1 Step -1
                            Try
                                sel.Add(mechs.Item(i))
                            Catch
                            End Try
                        Next
                        Try
                            sel.Delete()
                            msg.AppendLine("  cleared " & nMechs & " imported mechanism(s) from main_assembly.")
                        Catch exDel As Exception
                            msg.AppendLine("  ! could not delete imported mechanisms via Selection: " & exDel.Message)
                        End Try
                        sel.Clear()
                    End If
                End If
            End If
        Catch ex As Exception
            ' Silent or logged
        End Try

        ' 2) Remove any phantom Fix constraint on rb_base.1 in main_assembly
        Try
            Dim cons As Object = AssemblySwap.GetConstraints(root)
            If cons IsNot Nothing Then
                Dim nCons As Integer = 0
                Try
                    nCons = CInt(cons.Count)
                Catch
                    nCons = 0
                End Try

                Dim removedFixes As Integer = 0
                For i As Integer = nCons To 1 Step -1
                    Dim c As Object = Nothing
                    Try
                        c = cons.Item(i)
                    Catch
                        c = Nothing
                    End Try

                    If c IsNot Nothing Then
                        Dim isFix As Boolean = False
                        Try
                            ' catCstTypeFix = 0 in CATIA V5 CatConstraintType
                            If CInt(c.Type) = 0 Then isFix = True
                        Catch
                        End Try

                        Dim cName As String = SafeName(c)
                        If Not isFix AndAlso cName.StartsWith("Fix", StringComparison.OrdinalIgnoreCase) Then
                            isFix = True
                        End If

                        If isFix Then
                            Dim el1 As String = AssemblySwap.ElementName(c, 1)
                            If el1.IndexOf("rb_base", StringComparison.OrdinalIgnoreCase) >= 0 OrElse
                               el1.IndexOf("robot", StringComparison.OrdinalIgnoreCase) >= 0 Then
                                Try
                                    cons.Remove(i)
                                    removedFixes += 1
                                Catch
                                End Try
                            End If
                        End If
                    End If
                Next
                If removedFixes > 0 Then
                    msg.AppendLine("  removed " & removedFixes & " conflicting Fix constraint(s) on robot base in main_assembly.")
                End If
            End If
        Catch exCst As Exception
        End Try

        Return msg.ToString().TrimEnd()
    End Function


    Private Function SafeName(ByVal o As Object) As String
        Try
            Return CStr(o.Name)
        Catch
            Return "?"
        End Try
    End Function

End Module
