Option Explicit

Sub copypaste()
    ' Purpose: Demonstrates two different methods to duplicate data (Value Transfer vs. Clipboard Copy).
    ' Author: Ashlesh JH

    ' 1. Setup Sample Data
    Range("a1:a10") = "Ashlesh"

    ' Method 1: Direct Value Transfer (Fastest / Best Practice for Values)
    ' This transfers raw values directly without using the Windows Clipboard.
    ' It is significantly faster than Copy/Paste for large datasets.
    Range("b1:b10").Value = Range("a1:a10").Value

    ' Method 2: Standard Copy and Paste (Clipboard)
    ' This uses the Clipboard, which copies values, formulas, and formatting.
    Range("a1:a10").Copy
    
    ' Pastes everything (Values + Formats) into column D
    Range("d1:d10").PasteSpecial
    
    ' Clears the Clipboard (removes the "marching ants" border around the copied cells)
    Application.CutCopyMode = False
    
    ' Reset selection to a neutral cell
    Range("F1").Select
End Sub
