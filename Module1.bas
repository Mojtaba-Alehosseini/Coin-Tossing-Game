Attribute VB_Name = "Module1"
Sub coin_toss_game()
 
    head_pr = Cells(2, 7)
    tail_pr = Cells(2, 8)
    inter_no = Cells(2, 6)
    harry_no = 0
    tom_no = 0
    harry_win = 0
    Tom_win = 0
    For i = 1 To inter_no
        Cells(i + 1, 1) = i
        If Rnd() < head_pr Then
           Cells(i + 1, 2) = "        H"
           harry_no = harry_no + 1
           tom_no = tom_no - 1
        Else
           Cells(i + 1, 2) = "         T"
           harry_no = harry_no - 1
           tom_no = tom_no + 1
        End If
        Cells(i + 1, 3) = harry_no
        Cells(i + 1, 4) = tom_no
        If harry_no > 0 Then
           harry_win = harry_win + 1
        End If
        If tom_no > 0 Then
            Tom_win = Tom_win + 1
        End If
     Next i
     Cells(6, 6) = "Harry win!"
     Cells(6, 7) = harry_win
     Cells(7, 6) = "Tom_win"
     Cells(7, 7) = Tom_win
        
     If harry_win >= 95 Then
        x = "Harry win almost all!"
     Else
        If harry_win <= 5 Then
           x = "Harry Lost almost all!"
        Else
          If harry_win <= 55 And harry_win >= 45 Then
             x = "Almost equal!"
          Else
             x = "--"
          End If
       End If
    End If
    Cells(9, 6) = "Response:"
    Cells(9, 7) = x
       
End Sub
