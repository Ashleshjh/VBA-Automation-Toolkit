Option Explicit

Sub font_color()

    ' Populate the target range with sample text
    Range("a1:a10").Value = "Ashlesh"

    ' Change font color using built-in VBA color constants
    Range("a1:a10").Font.Color = vbWhite
    Range("a1:a10").Font.Color = vbBlack
    Range("a1:a10").Font.Color = vbYellow
    Range("a1:a10").Font.Color = vbRed
    Range("a1:a10").Font.Color = vbGreen
    Range("a1:a10").Font.Color = vbCyan
    Range("a1:a10").Font.Color = vbMagenta

    ' Change font color using Excel's legacy ColorIndex numbers (1 to 56)
    Range("a1:a10").Font.ColorIndex = 1   ' 1 = Black
    Range("a1:a10").Font.ColorIndex = 10  ' 10 = Green
    Range("a1:a10").Font.ColorIndex = 50  ' 50 = Sea Green
    
End Sub