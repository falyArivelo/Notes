Imports Guna.UI2.WinForms

Public Class NoteCard
    Inherits UserControl

    Public Property Note As Note
    Public Event NoteCliquee(note As Note)

    Private estSelectionnee As Boolean = False

    Private ReadOnly Couleur1 As Color = Color.MediumPurple
    Private ReadOnly Couleur2 As Color = Color.DeepSkyBlue
    Private ReadOnly Couleur3 As Color = Color.MediumSlateBlue
    Private ReadOnly Couleur4 As Color = Color.DodgerBlue


    ' Contrôles
    Private PanelCard As Guna2Panel
    Private Titre As Label
    Private Description As Label

    Public Sub New(note As Note)
        Me.Note = note
        Me.Margin = New Padding(2, 2, 2, 2) ' Gauche, Haut, Droite, Bas

        Me.AutoSize = False
        'Me.AutoSizeMode = AutoSizeMode.GrowAndShrink

        'Me.MinimumSize = New Size(60, 60)
        'PanelCard.MinimumSize = New Size(200, 60)

        ' === Guna2CustomGradientPanel ===
        PanelCard = New Guna2Panel With {
    .BorderRadius = 10,
    .BorderThickness = 0.5,
    .BorderColor = Color.WhiteSmoke,
    .BackColor = Color.Transparent,
    .Padding = New Padding(5),
    .AutoSize = True,
    .AutoSizeMode = AutoSizeMode.GrowAndShrink,
    .Dock = DockStyle.Fill
}
        'PanelCard.FillColor = Color.Red ' pour test
        'PanelCard.FillColor2 = Color.Orange

        ' === Titre ===
        Titre = New Label With {
            .Text = note.Titre,
            .Font = New Font("Segoe UI", 12, FontStyle.Bold),
            .ForeColor = Color.WhiteSmoke,
            .AutoSize = True,
            .Dock = DockStyle.Top,
            .BackColor = Color.Transparent
        }

        ' === Description (date ou résumé) ===
        Description = New Label With {
            .Text = note.DateCreation.ToShortDateString(),
            .Font = New Font("Segoe UI", 9),
            .ForeColor = Color.Gainsboro,
            .AutoSize = True,
            .Dock = DockStyle.Top,
            .BackColor = Color.Transparent
        }

        ' Événement de clic sur toute la carte
        AddHandler Me.Click, AddressOf NoteCard_Click
        AddHandler Titre.Click, AddressOf NoteCard_Click
        AddHandler Description.Click, AddressOf NoteCard_Click
        AddHandler PanelCard.Click, AddressOf NoteCard_Click

        ' Ajout des labels dans le panel
        PanelCard.Controls.Add(Description)
        PanelCard.Controls.Add(Titre)

        ' Ajout du panel dans le UserControl
        Me.Controls.Add(PanelCard)

    End Sub

    Private Sub NoteCard_Click(sender As Object, e As EventArgs)
        RaiseEvent NoteCliquee(Me.Note)
    End Sub

    ' Marque la note comme sélectionnée ou non (affiche bordure blanche)
    Public Sub SetSelection(selection As Boolean)
        estSelectionnee = selection

        If selection Then
            PanelCard.BorderColor = Color.White
            PanelCard.BorderThickness = 2

            ' Activer fond dégradé
            'PanelCard.FillColor = Color.YellowGreen
            PanelCard.FillColor = Color.FromArgb(100, 128, 128, 128)


        Else
            PanelCard.BorderColor = Color.WhiteSmoke
            PanelCard.BorderThickness = 0.5
            ' Fond transparent si non sélectionnée
            PanelCard.FillColor = Color.Transparent
        End If
    End Sub
End Class
