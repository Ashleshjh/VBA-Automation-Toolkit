Option Explicit

Sub coloringacell()
    ' Purpose: Demonstrates different methods to change cell background colors.
    ' Author: Ashlesh JH

    ' Method 1: The "Select then Act" approach
    ' First select the cell, then modify the ActiveCell property.
    ' Note: Useful for following user cursor movements, but generally slower.
    Range("a1").Select
    ActiveCell.Interior.Color = vbRed

    ' Method 2: The "Direct Reference" approach (Best Practice)
    ' Applies formatting directly to the target cell without moving the selection cursor.
    ' This method is faster and cleaner for automation.
    Range("a2").Interior.Color = vbGreen

    ' Move the selection cursor to cell B2
    Range("b2").Select
End Sub
