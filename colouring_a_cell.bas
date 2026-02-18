Option Explicit

Sub coloringacell()
Range("a1").Select
ActiveCell.Interior.Color = vbRed
Range("a2").Interior.Color = vbGreen
Range("b2").Select
End Sub
