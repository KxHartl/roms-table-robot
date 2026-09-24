Option Explicit On
Option Strict Off

' Read-only mirror of Rule R_Profile (master.CATPart, see PARAMETERS.md section 3).
'
' The app NEVER writes these values - R_Profile owns them. They live here only so the
' UI can compute input limits and draw the sketches before anything is sent to CATIA.
' If R_Profile changes, change this table too (it is the only duplicated knowledge in
' the app, and it is duplicated on purpose: limits must work while offline).
Public Class ProfileInfo

    Public ReadOnly Name As String          ' catalogue string, e.g. "45x60"
    Public ReadOnly SizeAlongWidth As Double  ' first number  -> lies along Y
    Public ReadOnly SizeAlongLength As Double ' second number -> lies along X
    Public ReadOnly GussetSize As Double
    Public ReadOnly Caption As String
    Public ReadOnly Hint As String

    Public Sub New(ByVal name_ As String, ByVal alongWidth As Double, ByVal alongLength As Double,
                   ByVal gusset As Double, ByVal caption_ As String, ByVal hint_ As String)
        Name = name_
        SizeAlongWidth = alongWidth
        SizeAlongLength = alongLength
        GussetSize = gusset
        Caption = caption_
        Hint = hint_
    End Sub

    ' Half of the section per direction - the model calls these
    ' profile_offset_width (Y) and profile_offset_length (X).
    Public ReadOnly Property OffsetWidth As Double
        Get
            Return SizeAlongWidth / 2.0
        End Get
    End Property

    Public ReadOnly Property OffsetLength As Double
        Get
            Return SizeAlongLength / 2.0
        End Get
    End Property

    ' Bracket for the RAIL-TO-RAIL joint in a corner - mirror of F_gusset_size_corner
    ' (master.CATPart, PARAMETERS.md section 5).
    '
    ' A corner bracket lies flat and lands on the SIDE faces of the two rails. The rails
    ' are mounted on edge, so those faces are SizeAlongLength tall - which is the other
    ' half-section from the one the post-to-rail bracket follows:
    '
    '     GussetSize       = SizeAlongWidth   (45 / 45 / 45 / 60 / 90)
    '     GussetSizeCorner = SizeAlongLength  (45 / 60 / 90 / 60 / 90)
    '
    ' They differ on 45x60 and 45x90, which is the whole reason the assembly carries two
    ' bracket sizes at once and GussetSwitcher swaps two groups.
    Public ReadOnly Property GussetSizeCorner As Double
        Get
            Return SizeAlongLength
        End Get
    End Property

End Class

Public Module ProfileCatalog

    Private ReadOnly _all As ProfileInfo() = {
        New ProfileInfo("45x45", 45, 45, 45, "45 x 45 mm",
                        "Lightest section. Smallest brackets, so the table can be shorter and lower."),
        New ProfileInfo("45x60", 45, 60, 45, "45 x 60 mm",
                        "Default. 45 mm across the width (Y), 60 mm along the length (X) - the 60 stands upright."),
        New ProfileInfo("45x90", 45, 90, 45, "45 x 90 mm",
                        "Wide face carries two bolt rows, so the bolts sit at 1/4 and 3/4 instead of on the centre line."),
        New ProfileInfo("60x60", 60, 60, 60, "60 x 60 mm",
                        "Stiffer square section. Needs 60 mm brackets, so the table grows accordingly."),
        New ProfileInfo("90x90", 90, 90, 90, "90 x 90 mm",
                        "Heaviest section. Brackets are 90 mm, which raises the minimum table size noticeably.")
    }

    Public ReadOnly Property All As ProfileInfo()
        Get
            Return _all
        End Get
    End Property

    Public Function Find(ByVal name As String) As ProfileInfo
        If name IsNot Nothing Then
            For Each p As ProfileInfo In _all
                If String.Equals(p.Name, name.Trim(), StringComparison.OrdinalIgnoreCase) Then Return p
            Next
        End If
        Return _all(1) ' 45x60 - the model default
    End Function

End Module
