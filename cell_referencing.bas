Option Explicit

Sub cellreferencing()

Cells(3, 4).Select
Cells(3, 4).Value = "Hi"
Cells(4, 4).Interior.Color = vbRed
ActiveCell.Value = 40
Range("a2").Interior.Color = vbRed

Sheets("cell referencing").Range("h1") = "HI"

End Sub


