Option Explicit

' PRIVATE SCOPE: Restricts this macro so it can only be executed by other subroutines 
' within this specific module. It intentionally hides the macro from the Excel 
' user interface (Alt+F8 menu) to prevent manual execution.
Private Sub ifelseifand()

    ' Check if the score is between 1 and 34
    If Range("A2").Value >= 1 And Range("A2").Value <= 34 Then
        Range("B2").Value = "Fail"
        
    ' Check if the score is between 35 and 60
    ElseIf Range("A2").Value >= 35 And Range("A2").Value <= 60 Then
        Range("B2").Value = "C Grade"
        
    ' Check if the score is between 61 and 80
    ElseIf Range("A2").Value >= 61 And Range("A2").Value <= 80 Then
        Range("B2").Value = "B Grade"
        
    ' Check if the score is between 81 and 100
    ElseIf Range("A2").Value >= 81 And Range("A2").Value <= 100 Then
        Range("B2").Value = "A Grade"
        
    ' Handle any input that falls outside the 1-100 range
    Else
        Range("B2").Value = "Invalid Input"
    End If

End Sub