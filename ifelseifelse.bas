Option Explicit

Private Sub ifelseifelse()

    ' Check if the score is 34 or below
    If Range("A2").Value <= 34 Then
        Range("B2").Value = "Fail"
        
    ' Check if the score is between 35 and 60
    ElseIf Range("A2").Value <= 60 Then
        Range("B2").Value = "C Grade"
        
    ' Check if the score is between 61 and 80
    ElseIf Range("A2").Value <= 80 Then
        Range("B2").Value = "B Grade"
        
    ' Handle any score 81 or above
    Else
        Range("B2").Value = "A Grade"
    End If

End Sub