Option Explicit On
Option Strict Off

Imports System.Text

' Swapping one component for another in an assembly, without losing the constraint that
' held it in place.
'
' This is the machinery that took five rounds to get right for the gussets on 16.08.2026.
' It is here, and not copied into each switcher, because every finding below is a property
' of the CATIA API rather than of gussets - the screws hit exactly the same walls:
'
'   Products has no Replace; ReferenceProduct is read-only
'       -> a swap has to be Remove + AddComponent + a rebuilt constraint.
'   Product.Constraints does not exist
'       -> constraints hang off Product.Connections("CATIAConstraints").
'   AddBiEltCst's enum must not be guessed
'       -> 1 is Offset on this release, not Coincidence. The type is COPIED from the
'          constraint being replaced.
'   CreateReferenceFromName returns an object for ANY string and never throws
'       -> "the call worked" proves nothing. The finished constraint's Status is compared
'          against the original's (which was connected by construction).
'   DisplayName gives the RESOLVED path, not the path through the publication
'       -> the path contains the old part's name, so it needs translating to the new one.
'          What that translation is differs per switcher, so it arrives as a delegate.
Public Module AssemblySwap

    ' What was found about one instance, read BEFORE anything is changed.
    Public Class SwapTarget
        Public InstanceName As String       ' name in the tree
        Public PartNumber As String         ' what it is now
        Public WantPartNumber As String     ' what it should become
        Public ConstraintName As String
        Public MasterRef As String          ' DisplayName of the far side
        Public OwnRef As String             ' DisplayName of our own side
        ' The far side's COM Reference itself. The master is never removed, so the object
        ' stays valid after the old constraint is gone - and an object beats a string,
        ' because DisplayName is a display name and need not be valid input for
        ' CreateReferenceFromName. That was the cause of components stacked at the origin.
        Public MasterRefObj As Object
        Public CstType As Integer = -1
        Public CstStatus As Integer = Integer.MinValue
        Public UsedRef As String = ""       ' which spelling finally connected
        ' Name of the publication this component mates on. Empty means "origin", which is
        ' what the gussets and screws use; the tool publishes TOOL_AXS instead.
        Public PubName As String = ""
    End Class

    ' The assembly's constraint collection.
    '
    ' NOT root.Constraints - that member does not exist on Product ("Public member
    ' 'Constraints' on type 'Product' not found", measured 16.08.2026). The direct property
    ' is kept as a fallback in case another release exposes it.
    Public Function GetConstraints(ByVal root As Object) As Object
        Try
            Dim c As Object = root.Connections("CATIAConstraints")
            If c IsNot Nothing Then Return c
        Catch
        End Try
        Try
            Return root.Constraints
        Catch
        End Try
        Return Nothing
    End Function

    ' The single constraint that positions the given instance, or Nothing when it cannot be
    ' identified - the caller then refuses to touch anything.
    Public Function ReadConstraintOf(ByVal root As Object, ByVal instName As String) As SwapTarget
        Dim cons As Object = GetConstraints(root)
        If cons Is Nothing Then Return Nothing

        Dim n As Integer
        Try
            n = CInt(cons.Count)
        Catch
            Return Nothing
        End Try

        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Dim cname As String = ""
            Try
                c = cons.Item(i)
                cname = CStr(c.Name)
            Catch
                c = Nothing
            End Try
            If c IsNot Nothing Then
                Dim a As String = ElementName(c, 1)
                Dim b As String = ElementName(c, 2)
                If a <> "" OrElse b <> "" Then
                    Dim mine As String = ""
                    Dim other As String = ""
                    Dim mineIsFirst As Boolean = False
                    If Mentions(a, instName) Then
                        mine = a : other = b : mineIsFirst = True
                    ElseIf Mentions(b, instName) Then
                        mine = b : other = a : mineIsFirst = False
                    End If

                    If mine <> "" AndAlso other <> "" Then
                        Dim t As New SwapTarget()
                        t.ConstraintName = cname
                        t.MasterRef = other
                        t.OwnRef = mine
                        t.MasterRefObj = ElementObject(c, If(mineIsFirst, 2, 1))
                        Try
                            t.CstType = CInt(c.Type)
                        Catch
                            t.CstType = -1
                        End Try
                        Try
                            t.CstStatus = CInt(c.Status)
                        Catch
                            t.CstStatus = Integer.MinValue
                        End Try
                        Return t
                    End If
                End If
            End If
        Next

        Return Nothing
    End Function

    ' Swap one instance. Throws with a readable message on failure; the caller decides
    ' whether to carry on (it should not - see the note in each switcher).
    Public Sub SwapOne(ByVal root As Object, ByVal t As SwapTarget, ByVal srcProd As Object,
                       ByVal translate As Func(Of String, String))
        RemoveConstraint(root, t.ConstraintName)
        RemoveInstance(root, t.InstanceName)

        Dim newInst As Object = root.Products.AddComponent(srcProd)
        ' CATIA names the instance after the new part number; put the old name back so the
        ' tree stays stable across swaps - and so the old reference string still describes it.
        Try
            newInst.Name = t.InstanceName
        Catch
        End Try

        AddCoincidence(root, t, translate)
    End Sub

    Public Sub RemoveConstraint(ByVal root As Object, ByVal cname As String)
        Dim cons As Object = GetConstraints(root)
        If cons Is Nothing Then Throw New Exception("constraint collection unreachable")
        Dim n As Integer = CInt(cons.Count)
        ' Backwards: removing shifts every index behind the current one.
        For i As Integer = n To 1 Step -1
            Dim nm As String = ""
            Try
                nm = CStr(cons.Item(i).Name)
            Catch
                nm = ""
            End Try
            If nm = cname Then
                cons.Remove(i)
                Return
            End If
        Next
    End Sub

    Public Sub RemoveInstance(ByVal root As Object, ByVal instName As String)
        Dim kids As Object = root.Products
        Dim n As Integer = CInt(kids.Count)
        For i As Integer = n To 1 Step -1
            Dim nm As String = ""
            Try
                nm = CStr(kids.Item(i).Name)
            Catch
                nm = ""
            End Try
            If nm = instName Then
                kids.Remove(i)
                Return
            End If
        Next
        Throw New Exception("instance '" & instName & "' not found when removing")
    End Sub

    ' Rebuild the constraint between the re-added instance and the master side the old one
    ' pointed at.
    '
    ' The master side reuses the ORIGINAL Reference object, so it needs no name round-trip
    ' at all. Only our own side has to be rebuilt, because that instance was deleted and
    ' re-added.
    Public Sub AddCoincidence(ByVal root As Object, ByVal t As SwapTarget,
                              ByVal translate As Func(Of String, String))
        Dim r2 As Object = t.MasterRefObj
        If r2 Is Nothing Then
            Try
                r2 = root.CreateReferenceFromName(t.MasterRef)
            Catch
                r2 = Nothing
            End Try
        End If
        If r2 Is Nothing Then
            Throw New Exception("master side reference lost (" & t.MasterRef & ")")
        End If

        Dim cons As Object = GetConstraints(root)
        If cons Is Nothing Then Throw New Exception("constraint collection unreachable")
        If t.CstType < 0 Then
            Throw New Exception("the original constraint type could not be read, " &
                                "so a replacement would be a guess")
        End If

        ' CreateReferenceFromName hands back an object for ANY string - it does not throw on
        ' nonsense - so "did the call succeed" proves nothing, and a constraint has to be
        ' proved connected after the fact.
        '
        ' NOT by comparing Status against the original's (measured 17.08.2026, and it cost a
        ' run): Status is UP-TO-DATE-NESS, not connectivity. A constraint built a moment ago
        ' has not been through an Update, so it can never report what a settled one reports -
        ' every candidate was rejected with "want status 0 ... status 1". The gussets only
        ' passed because the original happened to be out of date too.
        '
        ' What does prove it: read the finished constraint back and check that our side
        ' actually names this instance. A reference built from a nonsense string yields a
        ' constraint whose element does not resolve to it.
        Dim tried As New StringBuilder()
        For Each cand As String In OwnRefCandidates(root, t, translate)
            Dim r1 As Object = Nothing
            Try
                r1 = root.CreateReferenceFromName(cand)
            Catch
                r1 = Nothing
            End Try
            If r1 Is Nothing Then
                tried.Append(" [" & cand & ": no ref]")
                Continue For
            End If

            Dim made As Object = Nothing
            Try
                made = cons.AddBiEltCst(t.CstType, r1, r2)
            Catch ex As Exception
                tried.Append(" [" & cand & ": " & ex.Message & "]")
                Continue For
            End Try

            Dim why As String = RejectReason(made, t)
            If why = "" Then
                t.UsedRef = cand
                Return
            End If

            ' Attached to nothing recognisable - undo it and try the next spelling.
            tried.Append(" [" & cand & ": " & why & "]")
            RemoveConstraintObject(cons, made)
        Next

        Throw New Exception("no reference spelling connected;" & tried.ToString())
    End Sub

    ' Every spelling worth trying for the re-added instance's published axis system. The
    ' winner is normally the ORIGINAL constraint's own reference, translated to the new
    ' component - the resolved path carries the old part's name, so it does not survive a
    ' swap untouched.
    '
    ' The fallbacks matter more than they look. DisplayName gives the RESOLVED path even when
    ' the constraint was built on a publication, so t.OwnRef can read straight through the
    ' sub-assembly - measured 29.08.2026 on the tool:
    '     robot_mainAssembly/tool_1/adapter.1/!adapter/Axis System.1
    ' That spelling names TWO internal things (the instance adapter.1 and the feature
    ' Axis System.1) and reconnects after a swap only if both happen to be spelled the same
    ' inside the replacement. When they are not, the publication spelling is the way back -
    ' so it has to be the RIGHT publication name, which is why it is per-target and not
    ' hard-coded to "origin".
    Private Function OwnRefCandidates(ByVal root As Object, ByVal t As SwapTarget,
                                      ByVal translate As Func(Of String, String)) As List(Of String)
        Dim c As New List(Of String)

        If t.OwnRef <> "" Then
            Dim moved As String = t.OwnRef
            If translate IsNot Nothing Then
                Try
                    moved = translate(t.OwnRef)
                Catch
                    moved = t.OwnRef
                End Try
            End If
            If moved <> t.OwnRef Then c.Add(moved)
            c.Add(t.OwnRef)
        End If

        Dim pub As String = If(String.IsNullOrEmpty(t.PubName), PUB_ORIGIN, t.PubName)
        c.Add(t.InstanceName & "/!" & pub)
        c.Add(root.Name & "/" & t.InstanceName & "/!" & pub)
        c.Add(t.InstanceName & "/" & pub)

        ' The gussets and screws mate on "origin"; keep it as a last resort for anything that
        ' names a different publication, since an extra spelling costs one rejected attempt.
        If pub <> PUB_ORIGIN Then
            c.Add(t.InstanceName & "/!" & PUB_ORIGIN)
            c.Add(root.Name & "/" & t.InstanceName & "/!" & PUB_ORIGIN)
        End If
        Return c
    End Function

    ' The publication every positioned component exposes outwards. Switchers whose component
    ' publishes under another name set SwapTarget.PubName instead.
    Public Const PUB_ORIGIN As String = "origin"

    ' The one Status value that does mean "attached to nothing" (measured on the gussets,
    ' 16.08.2026). Every other value is some shade of up-to-date-ness and says nothing about
    ' whether the constraint found its geometry.
    Private Const STATUS_DISCONNECTED As Integer = 5

    ' Why a freshly built constraint should be thrown away, or "" when it is good.
    '
    ' The real test is the read-back: our side has to name the instance we just re-added.
    ' CreateReferenceFromName accepts any string, so a wrong spelling still produces a
    ' constraint - just one pointing at nothing this assembly knows.
    Private Function RejectReason(ByVal made As Object, ByVal t As SwapTarget) As String
        If made Is Nothing Then Return "no constraint"

        Dim st As Integer = StatusOf(made)
        If st = STATUS_DISCONNECTED Then Return "disconnected"

        Dim a As String = ElementName(made, 1)
        Dim b As String = ElementName(made, 2)
        If Not (Mentions(a, t.InstanceName) OrElse Mentions(b, t.InstanceName)) Then
            Return "does not resolve to " & t.InstanceName
        End If

        Return ""
    End Function

    Public Function StatusOf(ByVal c As Object) As Integer
        If c Is Nothing Then Return Integer.MinValue
        Try
            Return CInt(c.Status)
        Catch
            Return Integer.MinValue
        End Try
    End Function

    Public Sub RemoveConstraintObject(ByVal cons As Object, ByVal c As Object)
        Dim nm As String = ""
        Try
            nm = CStr(c.Name)
        Catch
            Return
        End Try
        Dim n As Integer = 0
        Try
            n = CInt(cons.Count)
        Catch
            Return
        End Try
        For i As Integer = n To 1 Step -1
            Dim cur As String = ""
            Try
                cur = CStr(cons.Item(i).Name)
            Catch
                cur = ""
            End Try
            If cur = nm Then
                Try
                    cons.Remove(i)
                Catch
                End Try
                Return
            End If
        Next
    End Sub

    Public Function ElementName(ByVal c As Object, ByVal idx As Integer) As String
        Dim r As Object = Nothing
        Try
            r = c.GetConstraintElement(idx)
        Catch
            Return ""
        End Try
        If r Is Nothing Then Return ""
        Try
            Return CStr(r.DisplayName)
        Catch
        End Try
        Try
            Return CStr(r.Name)
        Catch
        End Try
        Return ""
    End Function

    Public Function ElementObject(ByVal c As Object, ByVal idx As Integer) As Object
        Try
            Return c.GetConstraintElement(idx)
        Catch
            Return Nothing
        End Try
    End Function

    ' The instance name appears as a path segment, e.g. "table/gusset_LD_x.1/!origin".
    Public Function Mentions(ByVal refName As String, ByVal instName As String) As Boolean
        If refName = "" OrElse instName = "" Then Return False
        Return refName.IndexOf(instName, StringComparison.OrdinalIgnoreCase) >= 0
    End Function

    ' Self-diagnosis. When constraints cannot be matched, print what the assembly actually
    ' exposes - guessing which of three causes it is costs a round; this answers it.
    Public Function DumpConstraints(ByVal root As Object) As String
        Dim sb As New StringBuilder()
        Dim cons As Object = GetConstraints(root)
        If cons Is Nothing Then
            Return "The constraint collection is unreachable - neither" & vbCrLf &
                   "Connections(""CATIAConstraints"") nor .Constraints answered."
        End If

        Dim n As Integer = -1
        Try
            n = CInt(cons.Count)
        Catch ex As Exception
            Return "The constraint collection has no Count (" & ex.Message & ")."
        End Try

        sb.AppendLine("Assembly '" & SafeName(root) & "' reports " & n & " constraint(s).")
        If n <= 0 Then
            sb.AppendLine("None on the root product - they are probably owned by a")
            sb.AppendLine("sub-product, or the instances are positioned without constraints.")
            Return sb.ToString()
        End If

        Dim shown As Integer = 0
        For i As Integer = 1 To n
            Dim c As Object = Nothing
            Try
                c = cons.Item(i)
            Catch
                c = Nothing
            End Try
            If c IsNot Nothing Then
                sb.AppendLine("  [" & i & "] " & SafeName(c) &
                              "  el1=" & Quote(ElementName(c, 1)) &
                              "  el2=" & Quote(ElementName(c, 2)))
                shown += 1
                If shown >= 4 Then
                    sb.AppendLine("  ... (" & (n - shown) & " more)")
                    Exit For
                End If
            End If
        Next
        Return sb.ToString()
    End Function

    Public Function SafeName(ByVal o As Object) As String
        Try
            Return CStr(o.Name)
        Catch
            Return "?"
        End Try
    End Function

    Public Function Quote(ByVal s As String) As String
        If s = "" Then Return "<empty>"
        Return "'" & s & "'"
    End Function

    Public Function CountLines(ByVal sb As StringBuilder) As Integer
        Return sb.ToString().Split(New String() {vbCrLf}, StringSplitOptions.RemoveEmptyEntries).Length
    End Function

End Module
