Option Explicit On
Option Strict Off

' GSD zamjena aktivnog presjeka profila (port EnsureJoinHasSingleProfile iz CATVBA).
' U tijelu "profiles" nalaze se imenovani presjeci (po vrijednosti parametra "profile").
' Join "profile_active" u tijelu "profile_active" ima jedan ulaz koji zamjenjujemo.
'
' Pojednostavljenje (zadatak §1.6): gabarit presjeka je potpuno odreden s
' 2*profile_offset_width x 2*profile_offset_length, pa JEDNA parametarska skica
' "profile_simple" pokriva svih pet profila. Stariji plan (pet zasebnih "<profil>_simple")
' i dalje radi ako te povrsine postoje - trazimo prvo njih, pa zajednicku skicu.
' Vanjski gabarit je identican punom presjeku, pa zamjena ulaza Joina ne dira
' nijedan constraint ni link u sklopu.
'
' NAPOMENA 1: u .NET late bindingu TypeName(...) ne vraca CATIA tip ("HybridShapeExtract"...),
' pa element trazimo SAMO po imenu i oslanjamo se na Try/Catch oko CreateReferenceFromObject.
'
' NAPOMENA 2: skica u geometrijskom skupu NIJE u kolekciji HybridShapes. Stoji u
' HybridSketches (HybridBody.Sketches na R2020 puca s Err 438), a puni presjeci u GS-u
' "profiles" su HybridShapes. Zato prolazimo obje kolekcije, pa pod-skupove i tijela, a
' na kraju i Selection.Search po imenu - inace pojednostavljeni presjek nacrtan kao skica
' ostane nevidljiv i zamjena tiho padne natrag na puni presjek.
Public Module ProfileSwitcher

    Private Const GS_PROFILES As String = "profiles"
    Private Const GS_PROFILE_OUT As String = "profile_active"
    Private Const JOIN_NAME As String = "profile_active"
    Private Const SIMPLE_SUFFIX As String = "_simple"

    ' Jedna parametarska skica koja pokriva svih pet profila (TODO.md, faza 7.1).
    Public Const SHARED_SIMPLE As String = "profile_simple"

    Private Function EnsureHybridBody(ByVal part As Object, ByVal name As String) As Object
        Dim hbs As Object = part.HybridBodies
        Dim hb As Object = Nothing
        Try
            hb = hbs.Item(name)
        Catch
            hb = Nothing
        End Try
        If hb Is Nothing Then
            hb = hbs.Add()
            hb.Name = name
        End If
        Return hb
    End Function

    ' Kolekcija s tijela GS-a, ili Nothing ako je verzija CATIA-e nema.
    ' HybridBody NEMA .Sketches (Err 438 na R2020) - skice u geometrijskom skupu stoje
    ' u .HybridSketches. Body ih ima u .Sketches, pa probamo oboje.
    Private Function TryCollection(ByVal owner As Object, ByVal which As String) As Object
        Try
            Select Case which
                Case "HybridSketches" : Return owner.HybridSketches
                Case "Sketches" : Return owner.Sketches
                Case Else : Return owner.HybridShapes
            End Select
        Catch
            Return Nothing
        End Try
    End Function

    ' Prvi element zadanog imena u kolekciji -> Reference (ili Nothing).
    Private Function RefFromCollection(ByVal part As Object, ByVal coll As Object,
                                       ByVal nameWanted As String) As Object
        If coll Is Nothing Then Return Nothing

        Dim n As Integer = 0
        Try
            n = CInt(coll.Count)
        Catch
            Return Nothing
        End Try

        Dim i As Integer
        For i = 1 To n
            Dim item As Object = Nothing
            Dim itemName As String = ""
            Try
                item = coll.Item(i)
                itemName = CStr(item.Name)
            Catch
                itemName = ""
            End Try
            If item IsNot Nothing AndAlso
               String.Equals(itemName, nameWanted, StringComparison.OrdinalIgnoreCase) Then
                Try
                    Return part.CreateReferenceFromObject(item)
                Catch
                    ' nije referencabilan oblik; probaj sljedeci s istim imenom
                End Try
            End If
        Next
        Return Nothing
    End Function

    ' Jedan kontejner (GS ili Body): povrsine, skice, pa pod-skupovi.
    Private Function RefFromContainer(ByVal part As Object, ByVal owner As Object,
                                      ByVal nameWanted As String) As Object
        If owner Is Nothing Then Return Nothing

        Dim r As Object = RefFromCollection(part, TryCollection(owner, "HybridShapes"), nameWanted)
        If r IsNot Nothing Then Return r

        r = RefFromCollection(part, TryCollection(owner, "HybridSketches"), nameWanted)
        If r IsNot Nothing Then Return r

        r = RefFromCollection(part, TryCollection(owner, "Sketches"), nameWanted)
        If r IsNot Nothing Then Return r

        ' Pod-skupovi: presjek nacrtan u GS-u unutar GS-a inace ostaje nevidljiv.
        Dim subs As Object = Nothing
        Try
            subs = owner.HybridBodies
        Catch
            subs = Nothing
        End Try
        If subs Is Nothing Then Return Nothing

        Dim n As Integer = 0
        Try
            n = CInt(subs.Count)
        Catch
            Return Nothing
        End Try

        Dim i As Integer
        For i = 1 To n
            Dim child As Object = Nothing
            Try
                child = subs.Item(i)
            Catch
                child = Nothing
            End Try
            r = RefFromContainer(part, child, nameWanted)
            If r IsNot Nothing Then Return r
        Next
        Return Nothing
    End Function

    ' Zadnja mreza: CATIA-ina pretraga po imenu. Kolekcije ovise o verziji (vidi
    ' TryCollection), Selection.Search nade element bez obzira u kojem kontejneru sjedi.
    ' Selekciju ostavljamo praznom kakvu smo je i zatekli.
    Private Function RefFromSearch(ByVal part As Object, ByVal nameWanted As String) As Object
        Dim sel As Object = Nothing
        Try
            sel = part.Parent.Selection
        Catch
            Return Nothing
        End Try
        If sel Is Nothing Then Return Nothing

        Dim found As Object = Nothing
        Try
            sel.Clear()
            sel.Search("Name=" & nameWanted & ",all")
            Dim n As Integer = CInt(sel.Count)
            Dim i As Integer
            For i = 1 To n
                Dim v As Object = Nothing
                Dim vName As String = ""
                Try
                    v = sel.Item(i).Value
                    vName = CStr(v.Name)
                Catch
                    v = Nothing
                End Try
                If v IsNot Nothing AndAlso
                   String.Equals(vName, nameWanted, StringComparison.OrdinalIgnoreCase) Then
                    Try
                        found = part.CreateReferenceFromObject(v)
                        Exit For
                    Catch
                        found = Nothing
                    End Try
                End If
            Next
        Catch
            found = Nothing
        End Try

        Try
            sel.Clear()
        Catch
        End Try
        Return found
    End Function

    ' Presjek po imenu, bez obzira je li povrsina ili skica. Prvo GS "profiles"
    ' (rekurzivno), pa cijeli part - da presjek nacrtan malo sa strane ne obori zamjenu.
    Private Function GetProfileRef(ByVal part As Object, ByVal nameWanted As String) As Object
        Dim hbIn As Object = Nothing
        Try
            hbIn = part.HybridBodies.Item(GS_PROFILES)
        Catch
            hbIn = Nothing
        End Try

        Dim r As Object = RefFromContainer(part, hbIn, nameWanted)
        If r IsNot Nothing Then Return r

        ' Ostali geometrijski skupovi.
        r = RefFromContainer(part, part, nameWanted)
        If r IsNot Nothing Then Return r

        ' Tijela (skica u PartBody umjesto u GS-u).
        Dim bodies As Object = Nothing
        Try
            bodies = part.Bodies
        Catch
            bodies = Nothing
        End Try

        If bodies IsNot Nothing Then
            Dim n As Integer = 0
            Try
                n = CInt(bodies.Count)
            Catch
                n = 0
            End Try

            Dim i As Integer
            For i = 1 To n
                Dim b As Object = Nothing
                Try
                    b = bodies.Item(i)
                Catch
                    b = Nothing
                End Try
                r = RefFromContainer(part, b, nameWanted)
                If r IsNot Nothing Then Return r
            Next
        End If

        Return RefFromSearch(part, nameWanted)
    End Function

    ' Ime presjeka za trazeni profil. Kod pojednostavljenja idemo redom:
    '   1) "<profil>_simple"  - zasebna inacica po profilu, ako je itko nacrtao
    '   2) "profile_simple"   - zajednicka parametarska skica (ono sto je u modelu)
    '   3) puni presjek       - tiho, isti obrazac kao opcionalni parametri, da app radi
    '                           i dok pojednostavljenje nije domodelirano
    Private Function ResolveSectionName(ByVal part As Object, ByVal profName As String,
                                        ByVal simplified As Boolean) As String
        If Not simplified Then Return profName

        Dim perProfile As String = profName & SIMPLE_SUFFIX
        If GetProfileRef(part, perProfile) IsNot Nothing Then Return perProfile
        If GetProfileRef(part, SHARED_SIMPLE) IsNot Nothing Then Return SHARED_SIMPLE
        Return profName
    End Function

    ' Postavi Join "profile_active" da ima jedan ulaz = odabrani presjek.
    ' Vraca ime presjeka koje je stvarno spojeno, da pozivatelj moze javiti kad je
    ' pojednostavljenje trazeno a u modelu ga nema.
    Public Function EnsureJoinHasSingleProfile(ByVal part As Object, ByVal profName As String,
                                               Optional ByVal simplified As Boolean = False) As String
        Dim hbOut As Object = EnsureHybridBody(part, GS_PROFILE_OUT)

        Dim joinObj As Object = Nothing
        Try
            joinObj = hbOut.HybridShapes.Item(JOIN_NAME)
        Catch
            joinObj = Nothing
        End Try

        Dim wantedName As String = ResolveSectionName(part, profName, simplified)
        Dim tgtRef As Object = GetProfileRef(part, wantedName)
        If tgtRef Is Nothing Then
            Throw New Exception("Presjek '" & wantedName & "' nije pronaden u tijelu '" & GS_PROFILES & "'.")
        End If

        If joinObj Is Nothing Then
            ' Inicijalno kreiraj Join (one-input trik: dvaput isti ref pa makni drugi).
            Dim hsf As Object = part.HybridShapeFactory
            joinObj = hsf.AddNewJoin(tgtRef, tgtRef)
            joinObj.Name = JOIN_NAME
            hbOut.AppendHybridShape(joinObj)
            Try
                joinObj.RemoveElement(2)
            Catch
            End Try
            part.Update()
            Return wantedName
        End If

        ' Zamijeni prvi element i ukloni visak.
        joinObj.ReplaceElement(1, tgtRef)
        Dim n As Integer = CInt(joinObj.GetElementsSize())
        Dim k As Integer
        For k = n To 2 Step -1
            Try
                joinObj.RemoveElement(k)
            Catch
            End Try
        Next
        part.Update()
        Return wantedName
    End Function

End Module
