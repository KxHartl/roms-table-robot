Option Explicit On
Option Strict Off

Imports System.Text

' Putting the right screw in each bolted joint (TODO.md phase 6).
'
' Unlike the profile section and the gusset size, a screw length is not a parameter of
' anything in the assembly: every stock length is its own saved document, so that its Part
' Number can be the real designation ("DIN 7991 - M8x30") and the BOM comes out right on
' its own. The price of that decision is exactly this module - when the plate thickness
' moves, the component has to be exchanged.
'
' Rule R_Fasteners in the master says WHICH length each joint needs; this reads those
' numbers and makes the assembly agree.
'
' TWO ASSEMBLIES, because the joints do not all live in one. The three plate joints are
' inside table_mainAssembly; the four bolts that fasten the robot base go through the
' worktop into the base flange, and main_assembly is the only document where those two
' parts meet. Same machinery, two entry points, one private worker.
'
' WHICH INSTANCE BELONGS TO WHICH JOINT is the one question a Part Number cannot answer -
' two joints can legitimately want the same length, and then counting by length says
' nothing. So instances are matched by their INSTANCE NAME (screw_post, screw_rail_upper,
' screw_rail_lower, robot_bolt). Name them once when anchoring, and the copies a Reuse
' Pattern makes inherit the name with a numeric suffix.
'
' REUSE PATTERN: only the anchored instance carries a constraint of its own; the copies are
' placed by the pattern feature, which holds on to the anchor as its "initial product to
' duplicate". A swap is RemoveInstance + AddComponent - there is no Replace in the V5
' automation - so swapping the anchor deletes the very thing the pattern is duplicating and
' CATIA then refuses to update:
'
'     One of inputs of the Assembly Feature Pattern is unavailable:
'     Initial product to duplicate is missing.
'
' So a joint is swapped ALL OF IT or NONE of it. If even one instance of a joint has no
' constraint of its own, this module touches nothing in that joint and says why. Skipping
' is not timidity: half-swapped is a broken pattern plus three wrong screws, and that is
' worse than the mismatch it was trying to fix.
'
' Which is why screws are never patterned in this model, while the T-nuts and the washers
' and nuts under the robot base may be: those never change length, so nothing ever has to
' delete them.
Public Module ScrewSwitcher

    ' Part Number of a stock length, as the documents carry it.
    Public Function DesignationOf(ByVal family As FastenerFamily, ByVal lengthMm As Double) As String
        Return family.Designation(lengthMm)
    End Function

    ' The document holding that length.
    Public Function PathOf(ByVal cadRoot As String, ByVal family As FastenerFamily,
                           ByVal lengthMm As Double) As String
        Return AppConfig.CadFile(cadRoot, "04_fasteners", family.FileStem(lengthMm) & ".CATPart")
    End Function

    ' One joint: which family it takes, how its instances are named, and how many there
    ' should be.
    Private Class Joint
        Public ReadOnly Prefixes As String()   ' any of these starts an instance of this joint
        Public ReadOnly Caption As String
        Public ReadOnly Quantity As Integer
        Public ReadOnly Family As FastenerFamily
        Public WantLength As Double

        Public Sub New(ByVal prefixes_ As String(), ByVal caption_ As String, ByVal qty As Integer,
                       ByVal family_ As FastenerFamily)
            Prefixes = prefixes_
            Caption = caption_
            Quantity = qty
            Family = family_
        End Sub
    End Class

    ' The three plate joints, in the order the report lists them. Quantities are the hole
    ' counts of the patterns in the plates (TODO.md 6.4).
    '
    ' "screw_rail_upper" deliberately has no _x / _y in it: the worktop's rail screws sit in
    ' TWO patterns, split so the T-nuts can face along their own slot (TODO.md 6.4a). Both
    ' sets are the same screw at the same length, so the prefix spans them and the joint
    ' stays one joint - it is the nuts that care about the direction, not the screws.
    '
    ' Each joint accepts more than one prefix because the naming in the tree was done by
    ' hand: the short forms are there so a set called "post" is found as readily as
    ' "screw_post".
    Private Function PlateJoints(ByVal lenPost As Double, ByVal lenUpper As Double,
                                 ByVal lenLower As Double) As List(Of Joint)
        Dim f As FastenerFamily = FastenerCatalog.M8Countersunk
        Dim a As New Joint(New String() {"screw_post", "post"},
                           "worktop to posts", 4, f)
        Dim b As New Joint(New String() {"screw_rail_upper", "rail_upper"},
                           "worktop to upper rails (x + y)", 8, f)
        Dim c As New Joint(New String() {"screw_rail_lower", "rail_lower"},
                           "shelf to lower rails", 4, f)
        a.WantLength = lenPost
        b.WantLength = lenUpper
        c.WantLength = lenLower
        Return New List(Of Joint) From {a, b, c}
    End Function

    ' The robot base joint: four DIN 912 through the worktop into the base flange. In the
    ' model of 06.09.2026 they are named robot_bolt.1 ... .4 and hang off the publication
    ' robot_bolt_AXS in rb_base, reached through robot_mainAssembly.1.
    '
    ' The washers and nuts under the plate are NOT here. They never change length, so there
    ' is nothing to swap and no reason to risk their constraints - which is also why they
    ' may stay in their Reuse Pattern while the bolts may not.
    Private Function RobotMountJoints(ByVal lenRobotMount As Double) As List(Of Joint)
        Dim a As New Joint(New String() {"robot_bolt", "screw_robot_mount"},
                           "robot base to worktop", FastenerCatalog.RobotMountCount,
                           FastenerCatalog.M10CapHead)
        a.WantLength = lenRobotMount
        Return New List(Of Joint) From {a}
    End Function

    ' Main entry for the three plate joints. Returns the message for the user; throws only
    ' when the assembly is unreachable.
    Public Function EnsureScrewLengths(ByVal app As Object, ByVal tableDoc As Object,
                                       ByVal lenPost As Double, ByVal lenUpper As Double,
                                       ByVal lenLower As Double,
                                       ByVal cadRoot As String) As String
        Return EnsureLengths(app, tableDoc, PlateJoints(lenPost, lenUpper, lenLower),
                             cadRoot, "Screws",
                             "screw_post / screw_rail_upper_x / screw_rail_upper_y / screw_rail_lower")
    End Function

    ' Main entry for the robot base bolts. The document is main_assembly, not the table.
    Public Function EnsureRobotMountScrews(ByVal app As Object, ByVal mainDoc As Object,
                                           ByVal lenRobotMount As Double,
                                           ByVal cadRoot As String) As String
        Return EnsureLengths(app, mainDoc, RobotMountJoints(lenRobotMount),
                             cadRoot, "Robot mount",
                             "robot_bolt.1 ... .4")
    End Function


    ' The worker. Everything above it only decides which document and which joints.
    Private Function EnsureLengths(ByVal app As Object, ByVal assemblyDoc As Object,
                                   ByVal joints As List(Of Joint), ByVal cadRoot As String,
                                   ByVal label As String, ByVal namingHint As String) As String

        Dim root As Object = Nothing
        Try
            root = assemblyDoc.Product
        Catch
            root = Nothing
        End Try
        If root Is Nothing Then Throw New Exception("The assembly document has no Product.")

        Dim msg As New StringBuilder()
        Dim swapped As Integer = 0
        Dim already As Integer = 0
        Dim stuck As New StringBuilder()
        Dim unnamed As Integer = 0
        Dim lastFamily As FastenerFamily = joints(0).Family

        For Each j As Joint In joints
            lastFamily = j.Family

            If j.WantLength <= 0 Then
                msg.AppendLine("  " & j.Caption & ": the master has no length for it, skipped.")
                Continue For
            End If

            Dim want As String = DesignationOf(j.Family, j.WantLength)
            Dim wantTok As String = j.Family.Token(j.WantLength)
            Dim found As List(Of String) = InstancesOf(root, j)

            If found.Count = 0 Then
                unnamed += 1
                msg.AppendLine("  " & j.Caption & ": no instance named '" &
                               j.Prefixes(0) & "*'.")
                Continue For
            End If

            ' The count is not a blocker - only the anchor can be swapped anyway - but a
            ' joint short of screws is worth saying out loud while we are already counting.
            If found.Count <> j.Quantity Then
                msg.AppendLine("  " & j.Caption & ": " & found.Count &
                               " instance(s) carry the name, " & j.Quantity & " expected.")
            End If

            ' --- read every constraint before changing anything ------------------------
            Dim todo As New List(Of AssemblySwap.SwapTarget)
            Dim ok As Integer = 0
            Dim noConstraint As New List(Of String)

            For Each nm As String In found
                Dim pn As String = PartNumberOf(root, nm)
                ' Compared on the TOKEN, not on the whole part number. The document names
                ' agree on "M10x50" but a supplier's part number may carry the pitch or a
                ' material class after it, and none of that changes which screw it is.
                If HasToken(pn, wantTok) Then
                    ok += 1
                Else
                    Dim t As AssemblySwap.SwapTarget = AssemblySwap.ReadConstraintOf(root, nm)
                    If t Is Nothing Then
                        noConstraint.Add(nm & " (" & pn & ")")
                    Else
                        t.InstanceName = nm
                        t.PartNumber = pn
                        t.WantPartNumber = want
                        todo.Add(t)
                    End If
                End If
            Next

            already += ok

            If todo.Count = 0 AndAlso noConstraint.Count = 0 Then
                msg.AppendLine("  " & j.Caption & ": all " & ok & " already " & want & ".")
                Continue For
            End If

            ' All or nothing. An instance without its own constraint is a pattern copy, and
            ' the pattern is anchored on one of the instances we would be deleting. Swap the
            ' anchor and the pattern loses its input; the assembly then fails to update with
            ' "Initial product to duplicate is missing", and the three copies are still the
            ' old length on top of it.
            If noConstraint.Count > 0 Then
                msg.AppendLine("  " & j.Caption & ": NOT touched - " & noConstraint.Count &
                               " of " & found.Count & " instance(s) have no constraint of their own.")
                stuck.AppendLine("  " & j.Caption & ": left alone on purpose. These are pattern")
                stuck.AppendLine("  copies: " & String.Join(", ", noConstraint.ToArray()))
                stuck.AppendLine("  Swapping the anchor would delete what the pattern duplicates.")
                Continue For
            End If

            ' --- the replacement document ----------------------------------------------
            Dim path As String = PathOf(cadRoot, j.Family, j.WantLength)
            If Not IO.File.Exists(path) Then
                msg.AppendLine("  " & j.Caption & ": " & path & " does not exist, skipped.")
                Continue For
            End If

            Dim srcProd As Object = Nothing
            If todo.Count > 0 Then
                Dim srcDoc As Object = CatiaConnector.GetOrOpenDocument(app, path)
                Try
                    srcProd = srcDoc.Product
                Catch
                    srcProd = Nothing
                End Try
                If srcProd Is Nothing Then
                    msg.AppendLine("  " & j.Caption & ": " & path & " opened but has no Product.")
                    Continue For
                End If
            End If

            ' --- swap, one at a time ----------------------------------------------------
            Dim done As Integer = 0
            Dim failedHere As String = ""
            For Each t As AssemblySwap.SwapTarget In todo
                Try
                    AssemblySwap.SwapOne(root, t, srcProd, TranslatorFor(t, j.Family))
                    done += 1
                Catch ex As Exception
                    ' Stop at the FIRST failure of this joint. The component is already
                    ' exchanged by the time the constraint is rebuilt, so carrying on would
                    ' leave the rest unconstrained - which is what happened to the gussets
                    ' on 16.08. One broken instance is fixable by hand; eight is a revert.
                    failedHere = t.InstanceName & " - " & ex.Message
                    Exit For
                End Try
            Next

            swapped += done
            msg.AppendLine("  " & j.Caption & ": " & done & " -> " & want &
                           If(ok > 0, " (" & ok & " already correct)", "") & ".")
            ' Which spelling of the reference finally held. Printed once per joint: it is
            ' the thing to look at first if a screw ends up in the wrong place.
            If done > 0 AndAlso todo(0).UsedRef <> "" Then
                msg.AppendLine("      via " & todo(0).UsedRef)
            End If

            If failedHere <> "" Then
                stuck.AppendLine("  " & failedHere)
                stuck.AppendLine("  (stopped there; " & (todo.Count - done - 1) &
                                 " more in this joint were left alone)")
            End If

        Next

        ' --- the message ----------------------------------------------------------------
        Dim head As String
        If swapped = 0 AndAlso stuck.Length = 0 AndAlso unnamed = 0 Then
            head = label & ": already correct."
        Else
            head = label & ": " & swapped & " swapped."
        End If

        Dim outp As New StringBuilder()
        outp.AppendLine(head)
        outp.Append(msg.ToString())

        If unnamed > 0 Then
            outp.AppendLine()
            outp.AppendLine("Name the ANCHORED instance of each joint in the tree -")
            outp.AppendLine("  " & namingHint)
            outp.AppendLine("- and the app can tell them apart. Two joints may legitimately want the")
            outp.AppendLine("same length, so the part number alone cannot say which is which.")
            outp.AppendLine()
            outp.AppendLine("What the assembly calls its screws right now:")
            outp.AppendLine(ScrewNamesInTree(root, lastFamily))
        End If

        If stuck.Length > 0 Then
            outp.AppendLine()
            outp.AppendLine("These need a hand:")
            outp.Append(stuck.ToString())
            outp.AppendLine()
            outp.AppendLine("A joint left alone is a joint whose screws sit in a Reuse Pattern.")
            outp.AppendLine("Two ways out:")
            outp.AppendLine("  - by hand, this once: right-click the anchor > Components >")
            outp.AppendLine("    Replace Component, and replace EVERY instance of that reference.")
            outp.AppendLine("  - for good: delete the pattern and give each screw its own")
            outp.AppendLine("    constraint. A screw whose length follows the plate cannot live")
            outp.AppendLine("    in a pattern; the nuts and washers can, they never change.")
        End If

        Return outp.ToString().TrimEnd()
    End Function


    ' Instances whose tree name starts with any of the joint's prefixes. An instance is
    ' listed once however many prefixes it happens to match.
    Private Function InstancesOf(ByVal root As Object, ByVal j As Joint) As List(Of String)
        Dim list As New List(Of String)
        For Each nm As String In AllInstanceNames(root)
            For Each pre As String In j.Prefixes
                If nm.StartsWith(pre, StringComparison.OrdinalIgnoreCase) Then
                    If Not list.Contains(nm) Then list.Add(nm)
                End If
            Next
        Next
        Return list
    End Function

    Private Function AllInstanceNames(ByVal root As Object) As List(Of String)
        Dim list As New List(Of String)
        Dim kids As Object = root.Products
        Dim n As Integer = 0
        Try
            n = CInt(kids.Count)
        Catch
            n = 0
        End Try

        For i As Integer = 1 To n
            Dim nm As String = ""
            Try
                nm = CStr(kids.Item(i).Name)
            Catch
                nm = ""
            End Try
            If nm <> "" Then list.Add(nm)
        Next
        Return list
    End Function

    ' What the assembly actually calls its screws. Printed when a joint comes up empty, so
    ' one Apply answers "what are they named" instead of costing a round of guessing.
    Private Function ScrewNamesInTree(ByVal root As Object, ByVal family As FastenerFamily) As String
        Dim hits As New List(Of String)
        Dim kids As Object = root.Products
        Dim n As Integer = 0
        Try
            n = CInt(kids.Count)
        Catch
            n = 0
        End Try

        For i As Integer = 1 To n
            Dim nm As String = ""
            Dim pn As String = ""
            Try
                nm = CStr(kids.Item(i).Name)
                pn = CStr(kids.Item(i).PartNumber)
            Catch
            End Try
            If pn.IndexOf(family.TokenPrefix, StringComparison.OrdinalIgnoreCase) >= 0 Then
                Dim entry As String = nm & "  [" & pn & "]"
                If Not hits.Contains(entry) Then hits.Add(entry)
            End If
        Next

        If hits.Count = 0 Then Return "  (no component with " & family.TokenPrefix &
                                      " in its part number)"
        Return "  " & String.Join(vbCrLf & "  ", hits.ToArray())
    End Function

    Private Function PartNumberOf(ByVal root As Object, ByVal instName As String) As String
        Dim kids As Object = root.Products
        Dim n As Integer = 0
        Try
            n = CInt(kids.Count)
        Catch
            Return ""
        End Try
        For i As Integer = 1 To n
            Try
                If CStr(kids.Item(i).Name) = instName Then Return CStr(kids.Item(i).PartNumber)
            Catch
            End Try
        Next
        Return ""
    End Function

    ' Does this part number carry exactly this size? The digit test matters: without it
    ' "M8x1" would be found inside "M8x12", and the swap would decide a screw is already
    ' right when it is not.
    Private Function HasToken(ByVal partNumber As String, ByVal token As String) As Boolean
        If partNumber Is Nothing OrElse token = "" Then Return False
        Dim i As Integer = partNumber.IndexOf(token, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return False
        Dim after As Integer = i + token.Length
        If after < partNumber.Length AndAlso Char.IsDigit(partNumber(after)) Then Return False
        Return True
    End Function

    ' A path rewriter for one particular swap: M8x30 -> M8x40, M10x50 -> M10x60.
    '
    ' DisplayName hands back the RESOLVED path, which names the document rather than the
    ' publication - so it carries the old length and does not survive the swap untouched.
    ' One token covers every spelling that appears there, because the document name
    ' (DIN_7991_M8x30, DIN_912_M10x50_p1-5) and the part number (DIN 7991 - M8x30) share it.
    Private Function TranslatorFor(ByVal t As AssemblySwap.SwapTarget,
                                   ByVal family As FastenerFamily) As Func(Of String, String)
        Dim fromTok As String = LengthToken(t.PartNumber, family)
        Dim toTok As String = LengthToken(t.WantPartNumber, family)
        If fromTok = "" OrElse toTok = "" OrElse fromTok = toTok Then
            Return Function(p As String) p
        End If
        Return Function(p As String) p.Replace(fromTok, toTok)
    End Function

    ' "DIN 7991 - M8x30" -> "M8x30". Empty when the part number carries no length.
    Private Function LengthToken(ByVal partNumber As String, ByVal family As FastenerFamily) As String
        If partNumber Is Nothing Then Return ""
        Dim pre As String = family.TokenPrefix
        Dim i As Integer = partNumber.IndexOf(pre, StringComparison.OrdinalIgnoreCase)
        If i < 0 Then Return ""
        Dim j As Integer = i + pre.Length
        While j < partNumber.Length AndAlso Char.IsDigit(partNumber(j))
            j += 1
        End While
        If j = i + pre.Length Then Return ""
        Return partNumber.Substring(i, j - i)
    End Function

End Module
