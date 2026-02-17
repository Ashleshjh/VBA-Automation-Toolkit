 Option Explicit

Sub background_color()
    ' Purpose: Demonstrates changing cell background colors
    
    ' Method 1: Using VBA Color Constants (Limited to 8 basic colors)
      Range("a1:a10").Interior.Color = vbWhite
      Range("a1:a10").Interior.Color = vbBlack
      Range("a1:a10").Interior.Color = vbYellow
      Range("a1:a10").Interior.Color = vbRed
      Range("a1:a10").Interior.Color = vbGreen
      Range("a1:a10").Interior.Color = vbBlue
      Range("a1:a10").Interior.Color = vbCyan
      Range("a1:a10").Interior.Color = vbMagenta

      ' Method 2: Using ColorIndex (Legacy Excel 56-color palette)
        Range("a1:a10").Interior.ColorIndex = 1
        Range("a1:a10").Interior.ColorIndex = 10
        Range("a1:a10").Interior.ColorIndex = 49
        Range("a1:a10").Interior.ColorIndex = 55

End Sub





