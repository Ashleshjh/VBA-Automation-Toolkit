Option Explicit

Sub columnwidth()
    ' Purpose: Demonstrates how to adjust column widths using fixed values and AutoFit.
    ' Author: Ashlesh JH

    ' 1. Setting fixed column widths
    ' The value represents roughly the number of characters that can display in the cell.
    Range("a1").columnwidth = 25
    Range("a1").columnwidth = 15
    Range("a1").columnwidth = 5

    ' 2. Using AutoFit
    ' Automatically resizes the column to fit the longest text entry currently in that column.
    Range("a1").EntireColumn.AutoFit

    ' 3. Modifying width via the specific Cell's Column property
    ' This targets the column containing cell A3 (Column A) and sets the width.
    Range("a3").Columns.columnwidth = 35
    Range("a3").Columns.columnwidth = 5
    Range("a3").Columns.columnwidth = 20
End Sub
