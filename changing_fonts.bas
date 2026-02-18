Option Explicit

Sub changing_font()

Range("a1:a10") = "Ashlesh"

Range("a1:a10").Font.name = "Ariel"
Range("a1:a10").Font.Size = 20
Range("a1:a10").Font.Bold = True
Range("a1:a10").Font.Italic = True
Range("a1:a10").Font.Underline = True
Range("a1:a10").Font.Bold = False
Range("a1:a10").Font.Italic = False
Range("a1:a10").Font.Underline = False
Range("a1:a10").Font.Strikethrough = True
Range("a1:a10").Font.Strikethrough = False
End Sub

