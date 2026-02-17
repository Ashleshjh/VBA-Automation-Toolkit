Option Explicit

Sub ActivateAndSelect()
    ' Purpose: Demonstrates the difference between Selecting and Activating cells.
    ' Note: In professional automation, avoid .Select/.Activate to improve speed.
    
    ' Selects a single cell
    Range("A1").Select
    
    ' Selects a different single cell (moves cursor)
    Range("A3").Select
    
    ' Selects a group of cells
    Range("A1:D1").Select
    
    ' Activates a specific cell within a selection (or moves focus)
    Range("A4").Activate
    
    ' Activates a cell on a different range
    Range("B5").Activate
    
    ' Activates a range of cells
    Range("D1:E5").Activate
End Sub
