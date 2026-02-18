Option Explicit
Sub columnwidth()
Range("a1").columnwidth = 25
Range("a1").columnwidth = 15
Range("a1").columnwidth = 5
Range("a1").EntireColumn.AutoFit
Range("a3").Columns.columnwidth = 35
Range("a3").Columns.columnwidth = 5
Range("a3").Columns.columnwidth = 20
End Sub
