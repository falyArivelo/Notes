Imports Newtonsoft.Json
Imports System.IO

Public Class NoteManager
    Public Property Notes As New List(Of Note)

    Public Sub ChargerNotesDepuisDossier(Optional dossier As String = "Notes")
        Notes.Clear()

        If Not Directory.Exists(dossier) Then Directory.CreateDirectory(dossier)

        For Each fichier In Directory.GetFiles(dossier, "*.json")
            Try
                Dim json = File.ReadAllText(fichier)
                Dim note = JsonConvert.DeserializeObject(Of Note)(json)
                If note IsNot Nothing Then Notes.Add(note)
            Catch ex As Exception
                ' Gérer l'erreur de lecture si besoin
            End Try
        Next

        ' Trier : d'abord les épinglées, ensuite par date de création
        Notes = Notes.
            OrderByDescending(Function(n) n.Pin).
            ThenByDescending(Function(n) n.DateCreation).
            ToList()
    End Sub


End Class
