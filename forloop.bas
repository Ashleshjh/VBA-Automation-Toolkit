Option Explicit

Sub forloop()

    Dim i As Integer
    
    ' Loop through rows 2 to 6
    For i = 2 To 6
        
        ' Check if the value in Column A is greater than 34
        If Cells(i, 1).Value > 34 Then
            ' Mark as Pass in Column B
            Cells(i, 2).Value = "Pass"
        Else
            ' Otherwise, mark as Fail in Column B
            Cells(i, 2).Value = "Fail"
        End If
        
    Next i

End Sub