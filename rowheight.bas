Option Explicit

Sub rowheight()

    ' Set the row height for the row containing cell A2 to 25 points
    Range("a2").rowheight = 25
    
    ' Set the row height for the row containing cell B5 to 15 points
    Range("b5").rowheight = 15
    
    ' Set the row height for the row containing cell A3 to 5 points
    Range("a3").rowheight = 5
    
    ' Automatically adjust the row height of row 3 to fit its contents
    Range("a3").EntireRow.AutoFit
    
    ' Redundantly set the row height for row 2 back to 25 points
    Range("a2").Rows.rowheight = 25
    
    ' Immediately overwrite the previous command and set row 2 to 100 points
    Range("a2").Rows.rowheight = 100
    
    ' Immediately overwrite the previous command again and set row 2 to 40 points
    Range("a2").Rows.rowheight = 40

End Sub