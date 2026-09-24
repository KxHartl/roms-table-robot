Option Explicit On
Option Strict Off

Imports System.Globalization

' Pomocne numericke funkcije (port iz VBA: ToMm, ParseProfile, MapStandardProfile).
Public Module Num

    ' Normaliziraj korisnicki unos broja (zarez->tocka, makni razmake).
    Public Function Normalize(ByVal s As String) As String
        If s Is Nothing Then Return ""
        Return s.Replace(",", ".").Replace(" ", "")
    End Function

    Public Function IsNumber(ByVal s As String) As Boolean
        Dim d As Double
        Return Double.TryParse(Normalize(s), NumberStyles.Float, CultureInfo.InvariantCulture, d)
    End Function

    ' "12,5" -> "12.5mm" (uvijek tocka kao decimalni separator za CATIA-u).
    Public Function ToMm(ByVal valueText As String) As String
        Dim d As Double
        If Not Double.TryParse(Normalize(valueText), NumberStyles.Float, CultureInfo.InvariantCulture, d) Then
            Throw New FormatException("Nije broj: '" & valueText & "'")
        End If
        Return d.ToString(CultureInfo.InvariantCulture) & "mm"
    End Function

    Public Function MmString(ByVal d As Double) As String
        Return d.ToString(CultureInfo.InvariantCulture) & "mm"
    End Function

    ' "45x60" -> w=45, h=60 (vraca True ako uspije).
    Public Function ParseProfile(ByVal s As String, ByRef w As Double, ByRef h As Double) As Boolean
        If s Is Nothing Then Return False
        Dim t As String = s.ToLowerInvariant().Replace(" ", "")
        Dim parts() As String = t.Split("x"c)
        If parts.Length <> 2 Then Return False
        If Not Double.TryParse(parts(0), NumberStyles.Float, CultureInfo.InvariantCulture, w) Then Return False
        If Not Double.TryParse(parts(1), NumberStyles.Float, CultureInfo.InvariantCulture, h) Then Return False
        Return True
    End Function

End Module
