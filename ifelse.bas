Option Explicit

Sub ifelse()

    ' Check if the value in A2 is less than 35
    If Range("a2").Value < 35 Then
        ' Mark as Fail in B2
        Range("b2").Value = "Fail"
    Else
        ' Mark as Pass in B2
        Range("b2").Value = "Pass"
    End If

End Sub