Option Explicit On
Option Strict Off

' Puts the robot back into its home pose after the model has been rebuilt, and makes
' the DMU mechanism re-read the joint geometry.
'
' WHY THIS HAS TO EXIST  (measured 24.09.2026)
'
' Changing len_J2J3 / len_J3J5 moves the joint points inside rb_link2..6 (the formulas
' F_rb_linkN_* in robot_mainAssembly drive them from master\rb_z_J3 and friends). The
' links themselves are modelled in the base frame, so at home every link sits on the
' identity. But the mechanism keeps solving with the joint axes it had before the
' rebuild: after 300/320 -> 330/340 the tool missed P1 by 57 mm and P5 by 33 mm, with
' the joint points already correct. Putting every link back on the identity, the tool
' on the flange, and updating the mechanism brings both back to 0.000 mm.
'
' This is 03_macros\13_robot_set_home.CATScript, run inside CATIA through
' SystemService.Evaluate: Position.GetComponents / SetComponents take by-ref arrays,
' which late binding from .NET does not pass back reliably.
Public Module RobotHome

    Private Const SCRIPT As String =
        "Sub Mul(a, b, c)" & vbLf &
        " Dim i, j" & vbLf &
        " For j = 0 To 2 : For i = 0 To 2" & vbLf &
        "  c(j * 3 + i) = a(i) * b(j * 3) + a(3 + i) * b(j * 3 + 1) + a(6 + i) * b(j * 3 + 2)" & vbLf &
        " Next : Next" & vbLf &
        " For i = 0 To 2 : c(9 + i) = a(i) * b(9) + a(3 + i) * b(10) + a(6 + i) * b(11) + a(9 + i) : Next" & vbLf &
        "End Sub" & vbLf &
        "Sub Inv(a, c)" & vbLf &
        " Dim i, j" & vbLf &
        " For j = 0 To 2 : For i = 0 To 2 : c(j * 3 + i) = a(i * 3 + j) : Next : Next" & vbLf &
        " For i = 0 To 2 : c(9 + i) = -(a(i * 3) * a(9) + a(i * 3 + 1) * a(10) + a(i * 3 + 2) * a(11)) : Next" & vbLf &
        "End Sub" & vbLf &
        "Function CATMain(docName)" & vbLf &
        " Dim root, m, s, i, p, a(11), l6(11), il6(11), rel(11), ident" & vbLf &
        " On Error Resume Next" & vbLf &
        " Set root = CATIA.Documents.Item(docName).Product" & vbLf &
        " If Err.Number <> 0 Then CATMain = ""no document "" & docName : Exit Function" & vbLf &
        " Set m = root.GetTechnologicalObject(""Mechanisms"").Item(1)" & vbLf &
        " ident = Array(1, 0, 0, 0, 1, 0, 0, 0, 1, 0, 0, 0)" & vbLf &
        " root.Products.Item(""rb_link6.1"").Position.GetComponents l6" & vbLf &
        " If Err.Number <> 0 Then CATMain = ""no rb_link6.1"" : Exit Function" & vbLf &
        " Inv l6, il6" & vbLf &
        " For i = 1 To root.Products.Count" & vbLf &
        "  Set p = root.Products.Item(i)" & vbLf &
        "  p.Position.GetComponents a" & vbLf &
        "  If InStr(p.Name, ""rb_"") = 1 Then" & vbLf &
        "   p.Position.SetComponents ident" & vbLf &
        "  ElseIf InStr(p.Name, ""master"") = 0 Then" & vbLf &
        "   Mul il6, a, rel" & vbLf &
        "   p.Position.SetComponents rel" & vbLf &
        "  End If" & vbLf &
        "  If Err.Number <> 0 Then s = s & p.Name & "": "" & Err.Description & ""; "" : Err.Clear" & vbLf &
        " Next" & vbLf &
        " m.Update" & vbLf &
        " If Err.Number <> 0 Then s = s & ""mechanism: "" & Err.Description & ""; "" : Err.Clear" & vbLf &
        " root.Update" & vbLf &
        " If Err.Number <> 0 Then s = s & ""update: "" & Err.Description & ""; "" : Err.Clear" & vbLf &
        " CATMain = s" & vbLf &
        "End Function"

    ' Returns "" when the robot is home, otherwise what went wrong. Never throws: the
    ' parameters are already written and the model rebuilt, so a failure here is a line
    ' in the report, not a failed Apply.
    Public Function Reset(ByVal app As Object, ByVal robotDoc As Object) As String
        If app Is Nothing OrElse robotDoc Is Nothing Then Return ""
        Try
            Dim docName As String = CStr(robotDoc.Name)
            Dim r As Object = app.SystemService.Evaluate(SCRIPT, 0, "CATMain", New Object() {docName})
            Dim msg As String = If(r Is Nothing, "", CStr(r))
            If msg = "" Then Return ""
            Return "Robot home: " & msg
        Catch ex As Exception
            Return "Robot home: could not reset (" & ex.Message & "). Run " &
                   "03_macros\13_robot_set_home.CATScript fix by hand."
        End Try
    End Function

End Module
