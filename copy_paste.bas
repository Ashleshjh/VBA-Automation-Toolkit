Option Explicit

Sub copypaste()

Range("a1:a10") = "Ashlesh"
Range("b1:b10").Value = Range("a1:a10").Value

Range("a1:a10").Copy
Range("d1:d10").PasteSpecial
Application.CutCopyMode = False
Range("F1").Select

End Sub

