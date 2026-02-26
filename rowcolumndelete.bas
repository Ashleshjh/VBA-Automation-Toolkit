Option Explicit

Sub rowcolumndelete()

    ' Delete the entire column where cell A2 resides (Column A)
    Range("a2").EntireColumn.Delete
    
    ' Delete the entire row where cell D3 resides (Row 3)
    Range("d3").EntireRow.Delete
    
    ' Delete entire rows 1 through 3
    Range("a1:a3").EntireRow.Delete
    
    ' Delete entire columns where range B1:D1 resides (Columns B, C, D)
    Range("b1:d1").EntireColumn.Delete

End Sub