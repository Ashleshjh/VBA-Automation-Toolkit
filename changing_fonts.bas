Option Explicit

Sub changing_font()
    ' Purpose: Practice script to manipulate various Font properties (Name, Size, Styles).
    ' Author: Ashlesh JH
    
    ' 1. Assign a text value to the target range
    Range("a1:a10") = "Ashlesh"

    ' 2. Set Font Family and Size
    Range("a1:a10").Font.name = "Ariel" ' Sets the font face
    Range("a1:a10").Font.Size = 20      ' Increases font size to 20 points

    ' 3. Apply Formatting Styles (Turning properties ON)
    Range("a1:a10").Font.Bold = True        ' Applies Bold formatting
    Range("a1:a10").Font.Italic = True      ' Applies Italic formatting
    Range("a1:a10").Font.Underline = True   ' Applies Underline formatting

    ' 4. Remove Formatting Styles (Turning properties OFF)
    ' demonstrating how to toggle boolean properties back to False
    Range("a1:a10").Font.Bold = False       ' Removes Bold
    Range("a1:a10").Font.Italic = False     ' Removes Italic
    Range("a1:a10").Font.Underline = False  ' Removes Underline

    ' 5. Toggle Strikethrough
    Range("a1:a10").Font.Strikethrough = True  ' Applies Strikethrough line
    Range("a1:a10").Font.Strikethrough = False ' Removes Strikethrough line
End Sub
