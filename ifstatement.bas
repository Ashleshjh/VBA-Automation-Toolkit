Option Explicit

Sub ifstatement()

    ' If the value in A2 is greater than 35, execute the Pass marking on the same line
    If Range("a2").Value > 35 Then Range("b2").Value = "Pass"

End Sub