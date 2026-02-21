Sub deletecells()
    ' Deletes cell A1 (shifts remaining cells up by default)
    Range("a1").Delete
    
    ' Deletes the range A1:A10 (shifts remaining cells up by default)
    Range("a1:a10").Delete
    
    ' Deletes the entire rows containing the range A1:A10 (Rows 1 through 10)
    Range("a1:a10").EntireRow.Delete
    
    ' Deletes the entire columns containing the range A1:B1 (Columns A and B)
    Range("a1:b1").EntireColumn.Delete
End Sub