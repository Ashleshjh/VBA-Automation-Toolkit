Option Explicit

Sub CellReferencing()
    ' Purpose: Demonstrates different ways to reference cells (Range vs Cells)
    
    ' 1. Using Cells(Row, Column) - Best for loops
    Cells(3, 4).Select ' Selects D3 (Row 3, Col 4)
    Cells(3, 4).Value = "Hi"
    
    ' 2. Changing Properties
    Cells(4, 4).Interior.Color = vbRed ' Colors D4 Red
    
    ' 3. Using ActiveCell (The currently selected cell)
    ActiveCell.Value = 40
    
    ' 4. Using Standard A1 Notation
    Range("A2").Interior.Color = vbRed
    
    ' 5. Writing to a specific sheet without selecting it (Best Practice)
    Sheets("Sheet1").Range("H1").Value = "HI" 
End Sub
