Attribute VB_Name = "Module2"
Sub RoundedRectangle1_Click()
    head_pr = Cells(2, 10)
    tail_pr = Cells(2, 11)
    inter_no = Cells(2, 9)
    gen_no = Cells(2, 12)
    
    For f = 1 To gen_no
    
        Cells(f + 1, 6) = f
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
               Cells(i + 1, 2) = "        T"
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
        Cells(4, 9) = "Harry win"
        Cells(4, 10) = harry_win
        Cells(5, 9) = "Tom_win"
        Cells(5, 10) = Tom_win
        
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
        
    
       
       Cells(6, 9) = "Response:"
       Cells(6, 10) = x
       Cells(f + 1, 7) = x
       Next f
    
End Sub
