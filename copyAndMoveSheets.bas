Option Explicit

Sub copymovesheets()

    ' Create a duplicate of the hardcoded tab and place it after sheet8
    Sheets("cellreferencing").Copy after:=Sheets("sheet8")
    
    ' Move the original hardcoded tab to sit before sheet8
    Sheets("cellreferencing").Move before:=Sheets("sheet8")
    
    ' Create a duplicate of another hardcoded tab and place it after sheet8
    Sheets("addedwithVBArenamed").Copy after:=Sheets("sheet8")
    
    ' Move the original hardcoded tab to sit before sheet8
    Sheets("addedwithVBArenamed").Move before:=Sheets("sheet8")

End Sub