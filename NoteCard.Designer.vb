<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NoteCard
    Inherits System.Windows.Forms.UserControl

    'UserControl remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Titre = New Guna.UI2.WinForms.Guna2HtmlLabel()
        Description = New Guna.UI2.WinForms.Guna2HtmlLabel()
        SuspendLayout()
        ' 
        ' Titre
        ' 
        Titre.BackColor = Color.Transparent
        Titre.Location = New Point(6, 6)
        Titre.Name = "Titre"
        Titre.Size = New Size(25, 17)
        Titre.TabIndex = 0
        Titre.Text = "Title"
        ' 
        ' Description
        ' 
        Description.BackColor = Color.Transparent
        Description.Location = New Point(8, 29)
        Description.Name = "Description"
        Description.Size = New Size(63, 17)
        Description.TabIndex = 1
        Description.Text = "Description"
        ' 
        ' NoteCard
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        Controls.Add(Description)
        Controls.Add(Titre)
        Name = "NoteCard"
        Size = New Size(223, 53)
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents Titre As Guna.UI2.WinForms.Guna2HtmlLabel
    Friend WithEvents Description As Guna.UI2.WinForms.Guna2HtmlLabel

End Class
