Option Explicit
Sub paste_special()

    ' Populate cells A1 through A10 with a string value
    Range("a1:a10").Value = "Ashlesh"
    
    ' Copy the populated range to the clipboard
    Range("a1:a10").Copy
    
    ' Paste only the text values into column B
    Range("b1:b10").PasteSpecial xlPasteValues
    
    ' Match column B's width to column A
    Range("b1:b10").PasteSpecial xlPasteColumnWidths
    
    ' Apply column A's cell formatting (colors, borders, etc.) to column B
    Range("b1:b10").PasteSpecial xlPasteFormats
    
    ' Clear the clipboard memory to release Excel resources
    Application.CutCopyMode = False

End Sub