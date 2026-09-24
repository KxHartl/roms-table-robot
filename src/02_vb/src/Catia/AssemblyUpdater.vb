Option Explicit On
Option Strict Off

' Rekurzivni update sklopa (port UpdateAssemblyDeep / UpdateProductPartsRecursive).
' TypeName(...) nije pouzdan u .NET late bindingu, pa tip dokumenta odredujemo
' preko Try/Catch na svojstvima .Product / .Part.
Public Module AssemblyUpdater

    ' WITHDRAWN 30.08.2026: skipping a part document that had already been updated
    ' earlier in the same Apply.
    '
    ' The idea was sound on paper - main_assembly nests both other assemblies, so most
    ' parts were updated twice per Apply - and it saved a little of the update phase. But a
    ' rebuild came out wrong in the same session it was added, and an unproven optimisation
    ' sitting in the middle of a fault is worth less than the fault is worth understanding.
    ' The evidence says it was probably not the cause (a manual Update in CATIA did not fix
    ' the model either, so nothing was merely stale) - "probably" is not a reason to keep it
    ' while something is broken.
    '
    ' If it comes back, it comes back with a before/after 20_dump_state showing the same
    ' model. BeginPass is kept as a no-op so the call site keeps documenting the intent.
    Public Sub BeginPass()
    End Sub

    ' Azuriraj zadani product dokument (sklop) duboko.
    Public Sub UpdateProductDocument(ByVal prodDoc As Object)
        Dim prod As Object = Nothing
        Try
            prod = prodDoc.Product
        Catch
            prod = Nothing
        End Try
        If prod Is Nothing Then Return
        UpdateProductPartsRecursive(prod)
        prod.Update()
    End Sub

    ' Azuriraj ono sto je trenutno aktivno (fallback ako sklop nije eksplicitno poznat).
    Public Sub UpdateActive(ByVal app As Object)
        Dim doc As Object = app.ActiveDocument

        Dim prod As Object = Nothing
        Try
            prod = doc.Product
        Catch
            prod = Nothing
        End Try
        If prod IsNot Nothing Then
            UpdateProductPartsRecursive(prod)
            prod.Update()
            Return
        End If

        Dim part As Object = Nothing
        Try
            part = doc.Part
        Catch
            part = Nothing
        End Try
        If part IsNot Nothing Then part.Update()
    End Sub

    Private Sub UpdateProductPartsRecursive(ByVal prod As Object)
        Dim children As Object = prod.Products
        Dim count As Integer = CInt(children.Count)
        Dim i As Integer
        For i = 1 To count
            Dim inst As Object = children.Item(i)

            ' ReferenceProduct.Parent -> Document; ako ima .Part, pozovi Update.
            Dim refDoc As Object = Nothing
            Try
                refDoc = inst.ReferenceProduct.Parent
            Catch
                refDoc = Nothing
            End Try
            If refDoc IsNot Nothing Then
                Dim pd As Object = Nothing
                Try
                    pd = refDoc.Part
                Catch
                    pd = Nothing
                End Try
                If pd IsNot Nothing Then
                    Try
                        pd.Update()
                    Catch
                    End Try
                End If
            End If

            ' rekurzija u pod-sklopove
            Try
                If CInt(inst.Products.Count) > 0 Then
                    UpdateProductPartsRecursive(inst)
                End If
            Catch
            End Try
        Next
    End Sub

End Module
