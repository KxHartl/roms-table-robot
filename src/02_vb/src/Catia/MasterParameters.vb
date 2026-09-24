Option Explicit On
Option Strict Off

' Citanje/pisanje parametara iz master.CATPart (korijenski skup Parameters).
' Port logike iz CATVBA (GetParameterSafe / ToMm).
Public Class MasterParameters

    ' Imena parametara. Konvencija modela: length = X, width = Y, height = Z.
    Public Const P_LENGTH As String = "length"
    Public Const P_WIDTH As String = "width"
    Public Const P_HEIGHT As String = "height"
    Public Const P_PROFILE As String = "profile"
    Public Const P_UP_T As String = "upper_plate_thickness"
    Public Const P_LOW_T As String = "lower_plate_thickness"
    Public Const P_LOWER_M As String = "variant_lower_plate_mount"
    Public Const P_SIMP As String = "simplification"
    Public Const P_ROB_MNT As String = "variant_robot_mount"
    Public Const P_OFFX As String = "user_offset_x"
    Public Const P_OFFY As String = "user_offset_y"
    ' Polovica presjeka stupa PO SMJERU (ne po poziciji u stringu "45x60"):
    '   profile_offset_width  lezi po Y (sirina)  = prvi  broj / 2
    '   profile_offset_length lezi po X (duljina) = drugi broj / 2
    Public Const P_PROF_OFF_W As String = "profile_offset_width"
    Public Const P_PROF_OFF_L As String = "profile_offset_length"

    ' --- migracija naziva ---
    ' Model se preimenuje u jednoznacnu shemu vertical/width/length (gotovo).
    ' Dok to traje, app mora raditi i sa starim i s novim modelom, jer se .exe i .CATPart
    ' ne mijenjaju u istom trenutku. Zato se svako ime trazi prvo po novom nazivu, pa po
    ' starima iz ove tablice. Kad je preimenovanje gotovo, tablica se smije obrisati.
    Private Shared ReadOnly _legacyNames As New Dictionary(Of String, String()) From {
        {P_LENGTH, New String() {"lenght"}},
        {P_PROF_OFF_W, New String() {"profile_offset_x"}},
        {P_PROF_OFF_L, New String() {"profile_offset_y"}},
        {P_UP_T, New String() {"upperPlateThickness"}},
        {P_LOW_T, New String() {"lowerPlateThickness"}},
        {P_LOWER_M, New String() {"variantLowerPlateMount"}},
        {P_ROB_MNT, New String() {"variantRobotMount"}},
        {P_OFFX, New String() {"userOffsetX"}},
        {P_OFFY, New String() {"userOffsetY"}},
        {P_AXIS_OFF_X, New String() {"axisOffsetX"}},
        {P_AXIS_OFF_Y, New String() {"axisOffsetY"}},
        {P_PITCH_X, New String() {"h_robot_mount"}},
        {P_PITCH_Y, New String() {"v_robot_mount"}}
    }

    ' NAPOMENA: lower_plate_z i lower_plate_z_offset su UKINUTI u modelu 15.08.2026.
    ' Ravnina lower_plate_plane vec nosi height_lower_horizontal + profile_offset_length,
    ' a ploca se crta NA toj ravnini - pa ta dva parametra nisu imala potrosaca.
    ' Ne vracaj ih ovdje.

    ' Izvedeni parametri stola — samo za PRIKAZ (vodi ih Rule/formula u masteru,
    ' app ih nikad ne pise).
    ' Flange to gripping point, set by the rule R_Tool from measured values. The
    ' application READS it and never writes it: the master is the authority, and a second
    ' writer is how the .exe and the model came to hold two different lengths.
    Public Const P_TOOL_LEN As String = "tool_len"

    Public Const P_GUSSET_SIZE As String = "gusset_size"
    ' Kutnik ukruta<->ukruta. Vodi ga formula F_gusset_size_corner
    ' (= profile_offset_length * 2), pa se na 45x60 i 45x90 razlikuje od gusset_size.
    Public Const P_GUSSET_SIZE_CORNER As String = "gusset_size_corner"

    ' Driven by R_Fasteners: which stock screw each bolted joint needs. The app does not
    ' compute these into the model - it reads them back and makes the assembly agree.
    ' The first three are DIN 7991 in the plates; the fourth is the DIN 912 that bolts the
    ' robot base through the worktop, and it lives in main_assembly rather than in the
    ' table, because that is the only document where the base and the plate meet.
    Public Const P_SCREW_LEN_POST As String = "screw_len_post"
    Public Const P_SCREW_LEN_UPPER As String = "screw_len_upper"
    Public Const P_SCREW_LEN_LOWER As String = "screw_len_lower"
    Public Const P_SCREW_LEN_ROBOT_MOUNT As String = "screw_len_robot_mount"

    Public Const P_LOW_PLATE_LEN As String = "lower_plate_length"
    Public Const P_LOW_PLATE_WID As String = "lower_plate_width"

    ' Robot base raster, driven by R_RobotMountOffset. Read only, for the readout that
    ' puts the model's own numbers next to the app's - that is how a drifted rule or a
    ' misapplied sketch becomes visible instead of staying a surprise in the geometry.
    Public Const P_AXIS_OFF_X As String = "axis_offset_x"
    Public Const P_AXIS_OFF_Y As String = "axis_offset_y"
    Public Const P_PITCH_X As String = "robot_mount_pitch_x"
    Public Const P_PITCH_Y As String = "robot_mount_pitch_y"

    ' Raster baze robota. OD 17.08.2026. ove tri app PISE (t. 3 zadatka trazi promjenjiv
    ' raster). robot_mount_length/width su bile rucne konstante; R_RobotMountOffset iz njih
    ' i dalje vozi robot_mount_pitch_x/y po orijentaciji, pa se pravilo ne dira.
    ' robot_mount_hole_d je NOV - dok ne postoji u masteru, app ga uredno preskoci.
    Public Const P_MOUNT_LEN As String = "robot_mount_length"
    Public Const P_MOUNT_WID As String = "robot_mount_width"
    Public Const P_MOUNT_HOLE_D As String = "robot_mount_hole_d"

    ' Plate edge to the first bolt. Formula-driven (F_mount_edge_x/y) and the value the
    ' hole sketch is dimensioned against - a sketch dimension has no sign, so feeding it
    ' the centre-relative axis_offset mirrors the base back into the first quadrant.
    Public Const P_MOUNT_EDGE_X As String = "mount_edge_x"
    Public Const P_MOUNT_EDGE_Y As String = "mount_edge_y"

    ' Opcionalni (dodatni bodovi §6/§12) — pisi samo ako postoje u modelu.
    Public Const P_LEN_J2J3 As String = "len_J2J3"
    Public Const P_LEN_J3J5 As String = "len_J3J5"
    Public Const P_P1X As String = "P1_x"
    Public Const P_P1Y As String = "P1_y"
    Public Const P_P1Z As String = "P1_z"
    ' P3 is the pick station and the cube's own place; P6 the end point. Both became
    ' inputs on 29.08.2026, when the path stopped being laid out on a circle.
    Public Const P_P3X As String = "P3_x"
    Public Const P_P3Y As String = "P3_y"
    Public Const P_P3Z As String = "P3_z"
    Public Const P_P6X As String = "P6_x"
    Public Const P_P6Y As String = "P6_y"
    Public Const P_P6Z As String = "P6_z"
    Public Const P_P2P3 As String = "P2P3"
    Public Const P_P2P4 As String = "P2P4"

    ' P4P5 is NOT here on purpose: the task figure puts both stations at one height, so
    ' the model derives it (F_P4P5 = P2P3). Writing it would fight that formula.
    '
    ' path_radius and station_height are gone with the circle. They are left in the master
    ' without a consumer, like adapter_t, rather than deleted - see PARAMETERS.md 10.
    Public Const P_ADAPTER_T As String = "adapter_t"
    Public Const P_SIM_LEG As String = "sim_leg"
    ' Tool speed on the straight legs, mm/s. sim_leg times the joint ones - the two motion
    ' types are programmed with their own speeds, as on a real robot.
    Public Const P_SIM_SPEED As String = "sim_speed"
    ' The reconfiguration switch of task section 9. A STRING, "1" or "2" - R_Tool
    ' branches on it and ToolSwitcher exchanges the sub-assembly.
    Public Const P_OBJECT_TYPE As String = "object_type"

    Private ReadOnly _part As Object

    Public Sub New(ByVal part As Object)
        _part = part
    End Sub

    Public ReadOnly Property Part As Object
        Get
            Return _part
        End Get
    End Property

    ' --- dohvat parametra ---
    ' Nekad se prvo gledalo u skup "Table_Params". Taj skup u modelu nikad nije postojao
    ' (potvrdio master_inspect: svi parametri su ravno u korijenu), pa je grana samo tiho
    ' padala na FromAnywhere. Uklonjena 15.08.2026.

    Private Function FromAnywhere(ByVal paramName As String) As Object
        ' RootParameterSet.DirectParameters
        Try
            Dim rootSet As Object = _part.Parameters.RootParameterSet
            If rootSet IsNot Nothing Then
                Try
                    Dim p As Object = rootSet.DirectParameters.Item(paramName)
                    If p IsNot Nothing Then Return p
                Catch
                End Try
                ' svi pod-skupovi
                Try
                    Dim sets As Object = rootSet.ParameterSets
                    Dim i As Integer
                    For i = 1 To CInt(sets.Count)
                        Try
                            Dim ps As Object = sets.Item(i)
                            Dim p2 As Object = ps.DirectParameters.Item(paramName)
                            If p2 IsNot Nothing Then Return p2
                        Catch
                        End Try
                    Next
                Catch
                End Try
            End If
        Catch
        End Try
        ' globalna kolekcija
        Try
            Return _part.Parameters.Item(paramName)
        Catch
            Return Nothing
        End Try
    End Function

    ' Name -> parameter, remembered for the life of this object.
    '
    ' WHY. FromAnywhere is not a lookup, it is a SEARCH: it fetches RootParameterSet, tries
    ' DirectParameters.Item, then walks EVERY ParameterSet in the part trying each one, then
    ' falls back to a global Item - and for the dozen names that carry a legacy alias, the
    ' whole walk runs a second time. Every read, every write and every Exists paid for that
    ' again from scratch, and one Apply makes roughly 46 of them (TryWriteLengthReport alone
    ' did two per write: Exists, then WriteLength). Over COM, on a part with this many
    ' parameters, that is where the wait came from.
    '
    ' A parameter object does not move once found, and this class lives for one connection,
    ' so the answer is worth keeping. Misses are cached too - an absent optional parameter is
    ' asked about repeatedly by DetectMissingOptionals and by every TryOptional.
    Private ReadOnly _cache As New Dictionary(Of String, Object)(StringComparer.Ordinal)

    ' Trazi po novom imenu, pa po starima (migracija naziva — vidi _legacyNames).
    Public Function GetParam(ByVal name As String) As Object
        Dim hit As Object = Nothing
        If _cache.TryGetValue(name, hit) Then Return hit

        Dim p As Object = FindByExactName(name)
        If p Is Nothing Then
            Dim legacy As String() = Nothing
            If _legacyNames.TryGetValue(name, legacy) Then
                For Each alt As String In legacy
                    p = FindByExactName(alt)
                    If p IsNot Nothing Then Exit For
                Next
            End If
        End If

        ' Nothing is a legitimate cache entry: "this master does not have it" is exactly
        ' the answer DetectMissingOptionals wants, and it is the most expensive one to
        ' compute because it costs the full search plus every alias.
        _cache(name) = p
        Return p
    End Function


    ' Forgets everything. Only needed if parameters could be CREATED while this object is
    ' alive - a macro run in CATIA beside the application would do that - so the Reload
    ' button clears it rather than the app quietly serving a stale "does not exist".
    Public Sub ForgetCache()
        _cache.Clear()
    End Sub

    Private Function FindByExactName(ByVal name As String) As Object
        Return FromAnywhere(name)
    End Function

    Public Function Exists(ByVal name As String) As Boolean
        Return GetParam(name) IsNot Nothing
    End Function

    ' --- citanje ---

    Public Function ReadNumber(ByVal name As String) As String
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Return ""
        Try
            Return CStr(CDbl(p.Value))
        Catch
            Return ""
        End Try
    End Function

    Public Function ReadText(ByVal name As String) As String
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Return ""
        Try
            Return CStr(p.Value)
        Catch
            Return ""
        End Try
    End Function

    ' --- pisanje ---

    ' A write CATIA refuses is almost always a parameter owned by a formula or a rule -
    ' ValuateFromString / put_Value then fails with a message that does not name the
    ' parameter. Wrap every write so the caller learns WHICH one and the likely cause.
    Private Shared Function Blame(ByVal name As String, ByVal valueText As String,
                                  ByVal ex As Exception) As Exception
        Return New Exception(
            "Ne mogu upisati '" & name & "' = " & valueText & "." & vbCrLf &
            "Najcesci uzrok: taj parametar vodi formula ili pravilo (CATIA odbija vanjski upis)." & vbCrLf &
            "Provjeri u Tools -> Formula tko ga vodi." & vbCrLf & vbCrLf &
            "CATIA: " & ex.Message)
    End Function

    ' Duljinski parametar: prihvaca "12,5" ili "12.5" -> "12.5mm".
    Public Sub WriteLength(ByVal name As String, ByVal userText As String)
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Throw New Exception("Parametar nije pronaden: " & name)
        Dim mmText As String = Num.ToMm(userText)
        Try
            p.ValuateFromString(mmText)
        Catch ex As Exception
            Throw Blame(name, mmText, ex)
        End Try
    End Sub

    ' Duljina iz vec poznate mm vrijednosti.
    Public Sub WriteLengthMm(ByVal name As String, ByVal mm As Double)
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Throw New Exception("Parametar nije pronaden: " & name)
        Dim mmText As String = Num.MmString(mm)
        Try
            p.ValuateFromString(mmText)
        Catch ex As Exception
            Throw Blame(name, mmText, ex)
        End Try
    End Sub

    Public Sub WriteString(ByVal name As String, ByVal val As String)
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Throw New Exception("Parametar nije pronaden: " & name)
        Try
            p.Value = val
        Catch ex As Exception
            Throw Blame(name, "'" & val & "'", ex)
        End Try
    End Sub

    Public Sub WriteBool(ByVal name As String, ByVal val As Boolean)
        Dim p As Object = GetParam(name)
        If p Is Nothing Then Throw New Exception("Parametar nije pronaden: " & name)
        Try
            p.Value = val
        Catch ex As Exception
            Throw Blame(name, CStr(val), ex)
        End Try
    End Sub

    ' Opcionalni duljinski parametar.
    '   - ne postoji  -> False (stariji master, uredno preskoci)
    '   - postoji ali je read-only (vodi ga formula/pravilo) -> baci Blame s imenom,
    '     jer to NIJE benigno: len_J2J3 / P1_x moraju biti slobodni ulazi (t. 2.2 i t. 12),
    '     pa tiho preskakanje bi sakrilo da model ne moze primiti promjenu.
    Public Function TryWriteLength(ByVal name As String, ByVal userText As String) As Boolean
        If Not Exists(name) Then Return False
        WriteLength(name, userText)
        Return True
    End Function

    ' Kao TryWriteLength, ali read-only parametar prijavi u 'locked' umjesto da baci -
    ' *Apply* tako dovrsi ostale upise i na kraju pokaze POPIS zakljucanih, ne stane na prvom.
    Public Function TryWriteLengthReport(ByVal name As String, ByVal userText As String,
                                         ByVal locked As System.Text.StringBuilder) As Boolean
        ' GetParam, not Exists: Exists calls GetParam anyway, and calling both meant two
        ' lookups for one write. Cached now, but there is no reason to ask twice.
        If GetParam(name) Is Nothing Then Return False
        Try
            WriteLength(name, userText)
            Return True
        Catch ex As Exception
            locked.AppendLine("  - " & ex.Message.Replace(vbCrLf, " "))
            Return True
        End Try
    End Function

    ' UKLONJENO 15.08.2026: UpdateProfileOffsets(). Pisala je profile_offset_width/length
    ' iz stringa "45x60". Otkako te parametre vodi Rule R_Profile, taj upis je u najboljem
    ' slucaju suvisan, a u najgorem STETAN: ako pravilo ne radi, offseti bi i dalje
    ' izgledali tocno jer bi ih upisala app, dok bi gusset_size - koji postavlja samo
    ' pravilo - ostao zamrznut. Kvar bi se vidio tek na kutnicima.
    ' Sada, ako R_Profile ne radi, offseti odmah odu u krivo i kvar je ocit.

    Public Sub UpdatePart()
        _part.Update()
    End Sub

End Class
