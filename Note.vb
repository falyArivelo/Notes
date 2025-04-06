<Serializable>
Public Class Note
    Public Property Titre As String
    Public Property ContenuRTF As String ' Le texte RTF complet
    Public Property DateCreation As DateTime
    Public Property FichiersJoints As New List(Of Byte())
    Public Property Pin As Boolean
    Public Property DateAlarme As DateTime?
    Public Property Tags As New List(Of Tag)
End Class
