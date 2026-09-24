Option Explicit On
Option Strict Off

Imports System.IO

' Spajanje na CATIA-u i otvaranje dokumenata.
' Odluka: aplikacija SAMA otvara CATIA-u i datoteke (master + sklop).
Public Module CatiaConnector

    ' Spoji se na pokrenutu CATIA-u; ako je nema, pokreni novu.
    Public Function Connect() As Object
        Dim app As Object = Nothing
        Try
            app = GetObject(, "CATIA.Application")
        Catch
            app = Nothing
        End Try
        If app Is Nothing Then
            app = CreateObject("CATIA.Application")
        End If
        Try
            app.Visible = True
        Catch
            ' ignoriraj ako property nije dostupan
        End Try
        Return app
    End Function

    ' Vrati vec otvoreni dokument po imenu; ako nije otvoren, otvori ga s putanje.
    Public Function GetOrOpenDocument(ByVal app As Object, ByVal filePath As String) As Object
        Dim docs As Object = app.Documents
        Dim fileName As String = Path.GetFileName(filePath)

        Dim doc As Object = Nothing
        Try
            doc = docs.Item(fileName)
        Catch
            doc = Nothing
        End Try

        If doc Is Nothing Then
            If Not File.Exists(filePath) Then
                Throw New IO.FileNotFoundException("CATIA datoteka nije pronadena: " & filePath)
            End If
            doc = docs.Open(filePath)
        End If
        Return doc
    End Function

    ' Vrati dokument SAMO ako je vec otvoren; nikad ga ne otvara. Nothing inace.
    '
    ' Postoji zato da se smije pitati "je li robot sklop pri ruci?" bez cijene ucitavanja.
    ' GetOrOpenDocument bi ga otvorio, a to je nekoliko sekundi koje netko tko radi samo na
    ' stolu ne treba platiti.
    Public Function FindOpenDocument(ByVal app As Object, ByVal filePath As String) As Object
        If app Is Nothing Then Return Nothing
        Try
            Return app.Documents.Item(Path.GetFileName(filePath))
        Catch
            Return Nothing
        End Try
    End Function

    ' Podigni CATIA-in prozor ispred naseg.
    '
    ' Activate() ispod aktivira dokument UNUTAR CATIA-e, ali ne dize CATIA-u iznad ovog
    ' prozora - pa se simulacija odvozi iza ledja i izgleda kao da se nista nije dogodilo.
    ' Tako je i bilo prvi put: 22 poze odvozene, korisnik nije vidio nista.
    '
    ' Neuspjeh nije razlog da se odustane od simulacije, pa je sve u Try.
    Public Sub BringToFront()
        Try
            Dim procs() As Diagnostics.Process = Diagnostics.Process.GetProcessesByName("CNEXT")
            If procs.Length > 0 Then
                Microsoft.VisualBasic.Interaction.AppActivate(procs(0).Id)
            End If
        Catch
            ' Neki window manageri to ne dopustaju; simulacija svejedno radi.
        End Try
    End Sub

    ' Aktiviraj prozor dokumenta (da Update/ActiveDocument cilja pravi dokument).
    Public Sub Activate(ByVal doc As Object)
        Try
            doc.Activate()
        Catch
            ' neki dokumenti nemaju Activate; ignoriraj
        End Try
    End Sub

End Module
