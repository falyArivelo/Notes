Public Class NoteCard
    Public Property Note As Note

    Public Event NoteCliquee(note As Note)
    Private estSelectionnee As Boolean = False

    Private Sub NoteCard_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Titre.ForeColor = Color.WhiteSmoke
        Titre.Font = New Font("Segoe UI", 16, FontStyle.Bold)

        Description.ForeColor = Color.Gainsboro
        Description.Font = New Font("Segoe UI", 12, FontStyle.Regular)

    End Sub

    Public Sub Afficher(note As Note)
        Me.Note = note
        Titre.Text = note.Titre
        Description.Text = note.DateCreation.ToShortDateString()
    End Sub

    Private Sub NoteCard_Click(sender As Object, e As EventArgs) Handles MyBase.Click, Titre.Click, Description.Click
        RaiseEvent NoteCliquee(Me.Note)
    End Sub



    'Lorsque le Card est cliqué ou selectionner en cours card actif
    Public Sub SetSelection(estSelectionnee As Boolean)
        'If estSelectionnee Then
        '    'Me.BackColor = Color.White ' ou applique une bordure si tu as un Panel
        '    Me.BorderStyle = BorderStyle.FixedSingle
        'Else
        '    'Me.BackColor = SystemColors.Control
        '    Me.BorderStyle = BorderStyle.None
        'End If

        estSelectionnee = estSelectionnee
        Me.Invalidate() ' Force le redraw → Paint est relancé
    End Sub

    Private Sub NoteCard_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If estSelectionnee Then
            Dim couleurBordure As Color = Color.White
            Dim epaisseur As Integer = 3

            Using pen As New Pen(couleurBordure, epaisseur)
                pen.Alignment = Drawing2D.PenAlignment.Inset
                e.Graphics.DrawRectangle(pen, 0, 0, Me.Width - 1, Me.Height - 1)
            End Using
        End If
    End Sub



End Class
