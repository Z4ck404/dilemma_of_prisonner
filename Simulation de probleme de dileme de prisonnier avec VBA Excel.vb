'Tableau Fr pour faire le parcours des cellules excel'
Dim fr(19) As Variant
'Dim myarray As Variant'
'fr(0) = "B"'
'fr = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")'
Dim tab_matrice(19, 19, 1) 'la matrice qui va stocker l'etat de chaque case'
'la fonction qui vas récupérer la couleur d'une cellule excel'
Function GetColor(ByRef r As Range) As Integer
    GetColor = r.Interior.ColorIndex
End Function
Function SetColor(ByRef r As Range)
    r.Interior.ColorIndex = 1
End Function
'fill the data of the initial state of the game'
Sub Get_Data()
    fr(0) = "B"
    fr(1) = "C"
    fr(2) = "D"
    fr(3) = "E"
    fr(4) = "F"
    fr(5) = "G"
    fr(6) = "H"
    fr(7) = "I"
    fr(8) = "J"
    fr(9) = "K"
    fr(10) = "L"
    fr(11) = "M"
    fr(12) = "N"
    fr(13) = "O"
    fr(14) = "P"
    fr(15) = "Q"
    fr(16) = "R"
    fr(17) = "S"
    fr(18) = "T"
    fr(19) = "U"
    'fr = Array("B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U")'
    For i = 0 To 2
        For j = 0 To 2
            col = GetColor(Range(fr(i) & j + 2))
            If col = 3 Then
                tab_matrice(i, j, 0) = 1
            Else
                tab_matrice(i, j, 0) = 0
            End If
        Next
    Next
    'MsgBox tab_matrice(0, 0, 0)'
End Sub

'La fonction qui simule le jeu entre deux point'
Function test_play(ByVal x As Integer, ByVal y As Integer) As Integer
    If x = 1 Then
        If y = 1 Then
            test_play = 1
        Else
            If y = 0 Then
                test_play = -2
            End If
        End If
    Else
        If y = 1 Then
            test_play = 2
        Else
            test_play = 0
        End If
    End If
End Function
Sub iter_play()
    Call Get_Data
    For i = 0 To 19
        For j = 0 To 19
            If i = 0 And j = 0 Then
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                s8 = tab_matrice(i, j + 1, 0)
                s9 = tab_matrice(i + 1, j + 1, 0)
                scoref = test_play(s5, s5) + test_play(s5, s6) + test_play(s5, s9)
                'MsgBox scoref & " " & i & " " & j'
            ElseIf i > 0 And i < 19 And j = 0 Then
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                s7 = tab_matrice(i - 1, j + 1, 0)
                s8 = tab_matrice(i, j + 1, 0)
                s9 = tab_matrice(i + 1, j + 1, 0)
                score1 = test_play(s5, s4) + test_play(s5, s5) + test_play(s5, s6)
                score2 = test_play(s5, s7) + test_play(s5, s8) + test_play(s5, s9)
                scoref = score1 + score2
                'MsgBox scoref & " " & i & " " & j'
                
            ElseIf j > 0 And j < 19 And i = 19 Then
                s1 = tab_matrice(i - 1, j - 1, 0)
                s2 = tab_matrice(i, j - 1, 0)
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                s7 = tab_matrice(i - 1, j + 1, 0)
                s8 = tab_matrice(i, j + 1, 0)
                scoref = test_play(s5, s1) + test_play(s5, s2) + test_play(s5, s4) + test_play(s5, s5) + test_play(s5, s7) + test_play(s5, s8)
                'MsgBox scoref & " " & i & " " & j'

            ElseIf i = 0 And j > 0 And j <= 18 Then
                s2 = tab_matrice(i, j - 1, 0)
                s3 = tab_matrice(i + 1, j - 1, 0)
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                s8 = tab_matrice(i, j + 1, 0)
                s9 = tab_matrice(i + 1, j + 1, 0)
                scoref = test_play(s5, s2) + test_play(s5, s3) + test_play(s5, s5) + test_play(s5, s6) + test_play(s5, s8) + test_play(s5, s9)
                'MsgBox scoref & " " & i & " " & j & "here motherf**ker"'
            ElseIf i >= 1 And j >= 1 And i <= 18 And j <= 18 Then
            'le jeu pour les case qui sont pas dans les coins'
                s1 = tab_matrice(i - 1, j - 1, 0)
                s2 = tab_matrice(i, j - 1, 0)
                s3 = tab_matrice(i + 1, j - 1, 0)
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                s7 = tab_matrice(i - 1, j + 1, 0)
                s8 = tab_matrice(i, j + 1, 0)
                s9 = tab_matrice(i + 1, j + 1, 0)
                score1 = test_play(s5, s1) + test_play(s5, s2) + test_play(s5, s3)
                score2 = test_play(s5, s4) + test_play(s5, s5) + test_play(s5, s6)
                score3 = test_play(s5, s6) + test_play(s5, s8) + test_play(s5, s9)
                scoref = score1 + score2 + score3
                'MsgBox scoref & " " & i & " " & j'
            ElseIf j = 19 And i > 0 And i <= 18 Then
                s1 = tab_matrice(i - 1, j - 1, 0)
                s2 = tab_matrice(i, j - 1, 0)
                s3 = tab_matrice(i + 1, j - 1, 0)
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                score1 = test_play(s5, s1) + test_play(s5, s2) + test_play(s5, s3)
                score2 = test_play(s5, s4) + test_play(s5, s5) + test_play(s5, s6)
                scoref = score1 + score2
                'MsgBox scoref & " " & i & " " & j'
            ElseIf i = 0 And j = 19 Then
                s2 = tab_matrice(i, j - 1, 0)
                s3 = tab_matrice(i + 1, j - 1, 0)
                s5 = tab_matrice(i, j, 0)
                s6 = tab_matrice(i + 1, j, 0)
                scoref = test_play(s5, s2) + test_play(s5, s3) + test_play(s5, s5) + test_play(s5, s6)
                'MsgBox scoref & " " & i & " " & j & "here motherf"'
            ElseIf i = 19 And j = 19 Then
                s1 = tab_matrice(i - 1, j - 1, 0)
                s2 = tab_matrice(i, j - 1, 0)
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                scoref = test_play(s5, s1) + test_play(s5, s2) + test_play(s5, s4) + test_play(s5, s5)
                'MsgBox scoref & " " & i & " " & j'
            ElseIf i = 19 And j = 0 Then
                s4 = tab_matrice(i - 1, j, 0)
                s5 = tab_matrice(i, j, 0)
                s7 = tab_matrice(i - 1, j + 1, 0)
                s8 = tab_matrice(i, j + 1, 0)
                scoref = test_play(s5, s4) + test_play(s5, s5) + test_play(s5, s7) + test_play(s5, s8)
                'MsgBox scoref & " " & i & " " & j'
            End If
            If scoref > 0 Then
                    tab_matrice(i, j, 1) = "cop"
                    'MsgBox tab_matrice(i, j, 1) & " " & i & " " & j'
                    
            Else
                If scoref < 0 Then
                    tab_matrice(i, j, 1) = "def"
                    'MsgBox tab_matrice(i, j, 1) & " " & i & " " & j'
                    
                End If
            End If
        Next
    Next
End Sub
'Le macro main qui vas etre exceuté au click du button jouer'
Sub result_iter()
     Dim t As Integer
     'a chaque click du button , on peur changer le T '
     t = InputBox("Saisir la valuer de T ?", "La valeur de T")
     Call iter_play
     For i = 0 To 19
        For j = 0 To 19
            r = Range(fr(i) & i + 2)
            'Color = GetColor(r)'
            If tab_matrice(i, j, 1) = 1 Then
                If tab_matrice(i, j, 0) = 0 Then
                    r.Interior.ColorIndex = 3
                    tab_matrice(i, j, 0) = 1
                End If
            Else
                If tab_matrice(i, j, 0) = 1 Then
                    tab_matrice(i, j, 0) = 0
                    'r = Range(fr(i) & i + 2)'
                    Range(fr(i) & i + 2).Interior.ColorIndex = 2
                End If
            End If
            Next
     Next
'Range("A1").Interior.ColorIndex = 37'
End Sub





