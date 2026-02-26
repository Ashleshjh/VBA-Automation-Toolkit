Option Explicit

Sub rowcolumninsert()

    ' Insert a new column at Column A, shifting existing data to the right
    Range("a:a").Insert
    
    ' Insert a new row at Row 1, shifting existing data down
    Range("1:1").Insert
    
    ' Target cell A2 and insert an entire column, shifting data to the right
    Range("a2").EntireColumn.Insert
    
    ' Target cell C1 and insert an entire row, shifting data down
    Range("c1").EntireRow.Insert

End Sub