Option Explicit On
Option Strict Off

Imports System.Text

' Swapping the bracket size in the assembly (TODO.md phase 4.7 and 4.16).
'
' gusset_size (45/60/90) is driven by R_Profile and recomputes cleanly when the section
' changes - but that is only a number. It drives the height chain and the axis systems;
' WHICH component sits in the assembly is not something Knowledge can change, so the app
' does it, exactly as it does the profile section. Same channel, same reasoning (task SS6).
'
' TWO GROUPS, TWO SIZES (17.08.2026, TODO.md 4.12)
' The assembly holds 24 brackets in two families that can legitimately want DIFFERENT
' sizes at the same time:
'
'     post-to-rail   16  gusset_size         = 2 * profile_offset_width
'     rail-to-rail    8  gusset_size_corner  = 2 * profile_offset_length
'
' They differ on 45x60 and 45x90, because a rail is mounted on edge: the post joint lands
' on faces that are 2 * profile_offset_width across, the corner joint on the rails' SIDE
' faces, which are 2 * profile_offset_length tall.
'
' WHICH GROUP AN INSTANCE BELONGS TO is the one question a Part Number cannot answer - on
' a square section both groups carry the same one. So the corner brackets are recognised
' by their INSTANCE NAME (gusset_corner_lower_LD ...), the same mechanism ScrewSwitcher
' uses for the three plate joints. Everything else with a sub_gusset* part number is a
' post-to-rail bracket; the original sixteen keep their historical sub_gusset_45.N names
' and need no renaming.
'
' The mechanics of the swap - and the five findings it took to get them right - live in
' AssemblySwap, because none of them is about gussets. This module only knows which
' component it wants and how a bracket's reference path has to be rewritten.
'
' WHY IT IS SAFE HERE: every sub_gusset instance has EXACTLY ONE constraint, bound to the
' publication (!origin) which is addressed by string. The trap from roms-catscript-gotchas
' (scripted constraints breaking on internal edge IDs) does not apply - the publication has
' the same name in all three sizes. A manual Replace Component confirmed on 16.08.2026 that
' the model takes it.
'
' SAFETY RULE: read everything first, change nothing until all of it is understood. Half a
' swapped assembly is worse than an unswapped one. That check spans BOTH groups: an
' unreadable corner bracket stops the post brackets from being touched too.
Public Module GussetSwitcher

    ' Instances are recognised as brackets by a part number starting with this.
    Private Const PN_PREFIX As String = "sub_gusset"
    ' ... and as CORNER brackets by an instance name starting with this. Same spelling as
    ' the axis systems in the master and as PREFIX in 08_gusset_axes_corner.CATScript.
    Private Const CORNER_PREFIX As String = "gusset_corner"

    ' How many of each the finished table has (TODO.md 4.12). Not enforced - a count that
    ' is off is reported, because the swap itself is still correct for what is there.
    Private Const QTY_POST As Integer = 16
    Private Const QTY_CORNER As Integer = 8

    ' Target part number for a size. Empty when the size is not one of the three.
    Public Function TargetPartNumber(ByVal gussetSizeMm As Double) As String
        Dim s As Integer = CInt(Math.Round(gussetSizeMm))
        If s = 45 OrElse s = 60 OrElse s = 90 Then Return PN_PREFIX & "_" & s.ToString()
        Return ""
    End Function

    ' The CATProduct holding that size.
    Public Function TargetPath(ByVal cadRoot As String, ByVal gussetSizeMm As Double) As String
        Dim pn As String = TargetPartNumber(gussetSizeMm)
        If pn = "" Then Return ""
        Return AppConfig.CadFile(cadRoot, "01_table", pn & ".CATProduct")
    End Function

    ' One family of brackets: what it is called, what it wants, how many there should be.
    Private Class Group
        Public ReadOnly Caption As String
        Public ReadOnly Quantity As Integer
        Public ReadOnly WantSizeMm As Double
        Public ReadOnly WantPartNumber As String
        Public Found As Integer = 0
        Public Already As Integer = 0
        Public Swapped As Integer = 0

        Public Sub New(ByVal caption_ As String, ByVal qty As Integer, ByVal sizeMm As Double)
            Caption = caption_
            Quantity = qty
            WantSizeMm = sizeMm
            WantPartNumber = TargetPartNumber(sizeMm)
        End Sub
    End Class

    ' Main entry. Returns the message for the user; throws only when the assembly is
    ' unreachable.
    Public Function EnsureGussetSize(ByVal app As Object, ByVal tableDoc As Object,
                                     ByVal gussetSizeMm As Double,
                                     ByVal gussetSizeCornerMm As Double,
                                     ByVal cadRoot As String) As String

        Dim post As New Group("post-to-rail", QTY_POST, gussetSizeMm)
        Dim corner As New Group("rail-to-rail corners", QTY_CORNER, gussetSizeCornerMm)

        Dim bad As New StringBuilder()
        If post.WantPartNumber = "" Then
            bad.AppendLine("  gusset_size = " & gussetSizeMm & " mm is not one of 45/60/90.")
        End If
        If corner.WantPartNumber = "" Then
            bad.AppendLine("  gusset_size_corner = " & gussetSizeCornerMm &
                           " mm is not one of 45/60/90.")
        End If
        If bad.Length > 0 Then
            Return "Gussets: nothing to swap." & vbCrLf & bad.ToString().TrimEnd()
        End If

        Dim root As Object = Nothing
        Try
            root = tableDoc.Product
        Catch
            root = Nothing
        End Try
        If root Is Nothing Then Throw New Exception("The assembly document has no Product.")

        ' --- 1) read everything first ---------------------------------------------------
        Dim found As New List(Of AssemblySwap.SwapTarget)
        Dim owner As New List(Of Group)          ' parallel to found: which group each one is
        Dim problems As New StringBuilder()

        Dim kids As Object = root.Products
        Dim count As Integer = 0
        Try
            count = CInt(kids.Count)
        Catch
            count = 0
        End Try

        For i As Integer = 1 To count
            Dim pn As String = ""
            Dim nm As String = ""
            Try
                pn = CStr(kids.Item(i).PartNumber)
                nm = CStr(kids.Item(i).Name)
            Catch
                pn = ""
            End Try

            If Not pn.StartsWith(PN_PREFIX, StringComparison.OrdinalIgnoreCase) Then Continue For

            Dim g As Group = If(nm.StartsWith(CORNER_PREFIX, StringComparison.OrdinalIgnoreCase),
                                corner, post)
            g.Found += 1

            If String.Equals(pn, g.WantPartNumber, StringComparison.OrdinalIgnoreCase) Then
                g.Already += 1
                Continue For
            End If

            Dim f As AssemblySwap.SwapTarget = AssemblySwap.ReadConstraintOf(root, nm)
            If f Is Nothing Then
                problems.AppendLine("  " & nm & " (" & pn & ") - constraint not readable")
            Else
                f.InstanceName = nm
                f.PartNumber = pn
                f.WantPartNumber = g.WantPartNumber
                found.Add(f)
                owner.Add(g)
            End If
        Next

        If post.Found = 0 AndAlso corner.Found = 0 Then
            Return "Gussets: no " & PN_PREFIX & "* instance found in the assembly."
        End If

        ' Nothing is touched unless EVERY instance was understood - in either group.
        If problems.Length > 0 Then
            Return "Gussets NOT swapped - the assembly was left untouched." & vbCrLf &
                   "Could not read the constraint of " &
                   AssemblySwap.CountLines(problems) & " instance(s)." & vbCrLf & vbCrLf &
                   AssemblySwap.DumpConstraints(root) & vbCrLf &
                   "Swap them by hand: right-click an instance > Components >" & vbCrLf &
                   "Replace Component > the matching sub_gusset_*.CATProduct."
        End If

        If found.Count = 0 Then
            Return "Gussets: already correct." & vbCrLf & Tally(post) & vbCrLf & Tally(corner)
        End If

        ' --- 2) the replacement products --------------------------------------------------
        ' Cached by size: on a square section both groups want the same document, and
        ' opening it twice is a needless round trip through CATIA.
        Dim sources As New Dictionary(Of String, Object)(StringComparer.OrdinalIgnoreCase)
        For Each g As Group In New Group() {post, corner}
            If g.Found = 0 OrElse sources.ContainsKey(g.WantPartNumber) Then Continue For

            Dim path As String = TargetPath(cadRoot, g.WantSizeMm)
            If Not IO.File.Exists(path) Then
                Return "Gussets NOT swapped: " & path & " does not exist."
            End If

            Dim srcDoc As Object = CatiaConnector.GetOrOpenDocument(app, path)
            Dim srcProd As Object = Nothing
            Try
                srcProd = srcDoc.Product
            Catch
                srcProd = Nothing
            End Try
            If srcProd Is Nothing Then
                Return "Gussets NOT swapped: " & path & " opened but has no Product."
            End If
            sources(g.WantPartNumber) = srcProd
        Next

        ' --- 3) swap, one instance at a time ------------------------------------------------
        Dim failed As New StringBuilder()
        Dim done As Integer = 0

        For i As Integer = 0 To found.Count - 1
            Dim f As AssemblySwap.SwapTarget = found(i)
            Try
                AssemblySwap.SwapOne(root, f, sources(f.WantPartNumber), TranslatorFor(f))
                owner(i).Swapped += 1
                done += 1
            Catch ex As Exception
                ' Stop at the FIRST failure. The component is already swapped by the time the
                ' constraint is built, so carrying on would leave every gusset unconstrained -
                ' which is what happened on the run of 16.08. One broken instance is
                ' recoverable by hand; twenty-four is a revert.
                failed.AppendLine("  " & f.InstanceName & " - " & ex.Message)
                failed.AppendLine("  (stopped here; the remaining " &
                                  (found.Count - done - 1) & " were left alone)")
                Exit For
            End Try
        Next

        Dim msg As New StringBuilder()
        msg.AppendLine("Gussets: " & done & " swapped.")
        msg.AppendLine(Tally(post))
        msg.AppendLine(Tally(corner))

        ' Which spelling of the reference held. First thing to look at if a bracket ends up
        ' in the wrong place after a swap.
        If done > 0 AndAlso found(0).UsedRef <> "" Then
            msg.AppendLine("      via " & found(0).UsedRef)
        End If

        If corner.Found = 0 Then
            msg.AppendLine()
            msg.AppendLine("No corner bracket found. Name the eight rail-to-rail instances")
            msg.AppendLine("  gusset_corner_lower_LD/RD/RU/LU, gusset_corner_upper_...")
            msg.AppendLine("- otherwise they are swapped to gusset_size, which is the wrong")
            msg.AppendLine("size on 45x60 and 45x90.")
        End If

        If failed.Length > 0 Then
            msg.AppendLine()
            msg.AppendLine("These did NOT go through and need a hand:")
            msg.Append(failed.ToString())
        End If

        Return msg.ToString().TrimEnd()
    End Function


    ' One line per group: what it wanted, what it found, what it did.
    Private Function Tally(ByVal g As Group) As String
        Dim s As String = "  " & g.Caption & ": " & g.Found & " found -> " & g.WantPartNumber
        If g.Swapped > 0 Then s &= ", " & g.Swapped & " swapped"
        If g.Already > 0 Then s &= ", " & g.Already & " already correct"
        If g.Found <> g.Quantity Then s &= "  [expected " & g.Quantity & "]"
        Return s & "."
    End Function


    ' A path rewriter for one swap: GUSSET_45X45 -> GUSSET_60X60.
    '
    ' Measured 16.08.2026, a bracket's own reference reads
    '     table/sub_gusset_45.1/GUSSET_45X45.1/!GUSSET_45X45/Axis System.1
    ' so the publication is named after the gusset part rather than carrying one shared
    ' name, and every candidate that kept the old size came back disconnected (status 5
    ' against the original's 1) - which is exactly what the status check is there to catch.
    '
    ' The instance name itself (sub_gusset_45.1, gusset_corner_lower_LD) is NOT translated:
    ' it is restored after the swap, so it stays literally what it was.
    Private Function TranslatorFor(ByVal f As AssemblySwap.SwapTarget) As Func(Of String, String)
        Dim a As String = SizeOf(f.PartNumber)
        Dim b As String = SizeOf(f.WantPartNumber)
        If a = "" OrElse b = "" OrElse a = b Then Return Function(p As String) p

        Dim oldTok As String = "GUSSET_" & a & "X" & a
        Dim newTok As String = "GUSSET_" & b & "X" & b
        Return Function(p As String) p.Replace(oldTok, newTok)
    End Function

    ' "sub_gusset_45" -> "45"
    Private Function SizeOf(ByVal partNumber As String) As String
        If partNumber Is Nothing Then Return ""
        Dim i As Integer = partNumber.LastIndexOf("_"c)
        If i < 0 OrElse i = partNumber.Length - 1 Then Return ""
        Return partNumber.Substring(i + 1)
    End Function

End Module
