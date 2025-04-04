Public Module TextStyles

    ' Couleur utilisée pour mettre en évidence les boutons actifs
    Public ReadOnly activeColor As Color = ColorTranslator.FromHtml("#93B830")

    ' Optionnel : couleur inactive par défaut
    Public ReadOnly inactiveColor As Color = Color.Transparent

    Public ReadOnly Styles As New Dictionary(Of String, Single) From {
        {"title", 24},
        {"h1", 18},
        {"h2", 14},
        {"paragraph", 11}
    }
End Module
