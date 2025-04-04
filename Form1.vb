Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RichTextBox1.ScrollBars = RichTextBoxScrollBars.None
    End Sub


    Private Sub ApplyStyle(styleName As String)
        If RichTextBox1.SelectionFont IsNot Nothing Then
            Dim currentFont = RichTextBox1.SelectionFont

            Dim fontSize As Single

            If TextStyles.Styles.TryGetValue(styleName, fontSize) Then
                ' Appliquer le style tout en conservant la famille de police et le style (gras, italique…)
                RichTextBox1.SelectionFont = New Font(currentFont.FontFamily, fontSize, currentFont.Style)
            End If
        End If
    End Sub

    Private Sub ToggleStyle(style As FontStyle)
        If RichTextBox1.SelectionFont IsNot Nothing Then
            Dim currentFont = RichTextBox1.SelectionFont
            Dim newStyle As FontStyle = currentFont.Style Xor style ' toggle le style demandé
            RichTextBox1.SelectionFont = New Font(currentFont.FontFamily, currentFont.Size, newStyle)
        End If
    End Sub

    Private Sub UpdateBgButtonStates()
        Dim selectionFont = RichTextBox1.SelectionFont
        If selectionFont Is Nothing Then Return

        ButtonBold.BackColor = If(selectionFont.Bold, TextStyles.activeColor, TextStyles.inactiveColor)
        ButtonItalic.BackColor = If(selectionFont.Italic, TextStyles.activeColor, TextStyles.inactiveColor)
        ButtonUnderline.BackColor = If(selectionFont.Underline, TextStyles.activeColor, TextStyles.inactiveColor)

    End Sub


    Private Sub RichTextBox1_SelectionChanged(sender As Object, e As EventArgs) Handles RichTextBox1.SelectionChanged
        UpdateBgButtonStates()
    End Sub


    Private Sub Guna2ImageButton8_Click(sender As Object, e As EventArgs) Handles ButtonTitle.Click
        ApplyStyle("title")
    End Sub

    Private Sub Guna2ImageButton7_Click(sender As Object, e As EventArgs) Handles ButtonH1.Click
        ApplyStyle("h1")
    End Sub

    Private Sub Guna2ImageButton6_Click(sender As Object, e As EventArgs) Handles ButtonH2.Click
        ApplyStyle("h2")
    End Sub

    Private Sub Guna2ImageButton5_Click(sender As Object, e As EventArgs) Handles ButtonBold.Click
        ToggleStyle(FontStyle.Bold)
        UpdateBgButtonStates()
    End Sub

    Private Sub Guna2ImageButton4_Click(sender As Object, e As EventArgs) Handles ButtonUnderline.Click
        ToggleStyle(FontStyle.Underline)
        UpdateBgButtonStates()
    End Sub

    Private Sub Guna2ImageButton3_Click(sender As Object, e As EventArgs) Handles ButtonItalic.Click
        ToggleStyle(FontStyle.Italic)
        UpdateBgButtonStates()
    End Sub

    Private Sub Guna2ImageButton1_Click(sender As Object, e As EventArgs) Handles ButtonCote.Click

    End Sub

    Private Sub Guna2ImageButton2_Click(sender As Object, e As EventArgs) Handles ButtonAddFile.Click

    End Sub

    Private Sub ButtonCheckBox_Click(sender As Object, e As EventArgs) Handles ButtonCheckBox.Click
        Dim startLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)
        Dim endLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart + RichTextBox1.SelectionLength)

        For line = startLine To endLine
            Dim firstChar = RichTextBox1.GetFirstCharIndexFromLine(line)
            If firstChar >= 0 Then
                RichTextBox1.Select(firstChar, 0)
                RichTextBox1.SelectedText = "☐ " ' ou "[ ] " selon ton style
            End If
        Next
    End Sub

    Private Sub ToggleCheckboxAtCursor()
        Dim lineIndex = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)
        Dim lineStart = RichTextBox1.GetFirstCharIndexFromLine(lineIndex)
        Dim lineText = RichTextBox1.Lines(lineIndex)

        If lineText.StartsWith("☐") Then
            RichTextBox1.Select(lineStart, 2)
            RichTextBox1.SelectedText = "☑"
        ElseIf lineText.StartsWith("☑") Then
            RichTextBox1.Select(lineStart, 2)
            RichTextBox1.SelectedText = "☐"
        End If
    End Sub

    Private Sub ToggleCheckboxAtPosition(charIndex As Integer)
        Dim lineIndex = RichTextBox1.GetLineFromCharIndex(charIndex)
        Dim lineStart = RichTextBox1.GetFirstCharIndexFromLine(lineIndex)

        If lineStart >= 0 AndAlso RichTextBox1.Lines.Length > lineIndex Then
            RichTextBox1.Select(lineStart, 2)
            Dim current = RichTextBox1.SelectedText

            If current.StartsWith("☐") Then
                RichTextBox1.SelectedText = "☑ "
            ElseIf current.StartsWith("☑") Then
                RichTextBox1.SelectedText = "☐ "
            End If
        End If
    End Sub


    Private Sub RichTextBox1_MouseDown(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseDown
        Dim charIndex As Integer = RichTextBox1.GetCharIndexFromPosition(e.Location)

        If charIndex >= 0 AndAlso charIndex < RichTextBox1.TextLength Then
            Dim clickedChar As Char = RichTextBox1.Text(charIndex)

            If clickedChar = "☐"c OrElse clickedChar = "☑"c Then
                ToggleCheckboxAtPosition(charIndex)
            End If
        End If
    End Sub


    Private Sub RichTextBox1_MouseMove(sender As Object, e As MouseEventArgs) Handles RichTextBox1.MouseMove
        Dim charIndex As Integer = RichTextBox1.GetCharIndexFromPosition(e.Location)

        If charIndex >= 0 AndAlso charIndex < RichTextBox1.TextLength Then
            Dim hoveredChar As Char = RichTextBox1.Text(charIndex)

            If hoveredChar = "☐"c OrElse hoveredChar = "☑"c Then
                RichTextBox1.Cursor = Cursors.Hand ' Main/pointer
            Else
                RichTextBox1.Cursor = Cursors.IBeam ' Curseur texte standard
            End If
        End If
    End Sub

    Private Sub ButtonUL_Click(sender As Object, e As EventArgs) Handles ButtonUL.Click
        ToggleUnorderedList()
    End Sub

    Private Sub ButtonOL_Click(sender As Object, e As EventArgs) Handles ButtonOL.Click
        ToggleOrderedList()
    End Sub

    Private Sub ToggleUnorderedList()
        Dim startLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)
        Dim endLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart + RichTextBox1.SelectionLength)

        Dim allLinesAreBulleted = True

        ' Vérifie si toutes les lignes sont déjà avec puce
        For i = startLine To endLine
            If Not RichTextBox1.Lines(i).TrimStart().StartsWith("•") Then
                allLinesAreBulleted = False
                Exit For
            End If
        Next

        For i = startLine To endLine
            Dim lineStart = RichTextBox1.GetFirstCharIndexFromLine(i)
            Dim lineText = RichTextBox1.Lines(i)

            RichTextBox1.Select(lineStart, lineText.Length)

            If allLinesAreBulleted Then
                ' Supprimer la puce
                If lineText.TrimStart().StartsWith("•") Then
                    Dim index = lineText.IndexOf("•")
                    RichTextBox1.SelectedText = lineText.Remove(index, 2) ' "• " = 2 caractères
                End If
            Else
                ' Ajouter la puce
                RichTextBox1.SelectedText = "• " & lineText
            End If
        Next
    End Sub


    Private Sub ToggleOrderedList()
        Dim startLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart)
        Dim endLine = RichTextBox1.GetLineFromCharIndex(RichTextBox1.SelectionStart + RichTextBox1.SelectionLength)

        Dim allLinesAreNumbered = True

        ' Vérifie si toutes les lignes commencent par un numéro
        For i = startLine To endLine
            Dim line = RichTextBox1.Lines(i).TrimStart()
            If Not line Like "#.*" Then
                allLinesAreNumbered = False
                Exit For
            End If
        Next

        Dim number = 1
        For i = startLine To endLine
            Dim lineStart = RichTextBox1.GetFirstCharIndexFromLine(i)
            Dim lineText = RichTextBox1.Lines(i)

            RichTextBox1.Select(lineStart, lineText.Length)

            If allLinesAreNumbered AndAlso lineText.Trim() Like "#.*" Then
                ' Retirer la numérotation (ex: "1. ")
                Dim parts = lineText.Split(New Char() {" "c}, 2)
                If parts.Length = 2 Then
                    RichTextBox1.SelectedText = parts(1)
                End If
            Else
                ' Ajouter numérotation
                RichTextBox1.SelectedText = number.ToString() & ". " & lineText
            End If

            number += 1
        Next
    End Sub

End Class
