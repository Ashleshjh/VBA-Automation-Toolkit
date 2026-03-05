' Suppose I have 20 sheets. I have to access like example sheet5 and a particular cell in that sheet.
' With a module we can write the code in any active sheets. We can refer the cell into a particular sheet not in an active sheet.
Option Explicit

Sub sheetreferencing()

    ' Write to the first physical tab using the Worksheets collection
    Worksheets(1).Range("a1") = "VBA"
    
    ' Write to a hardcoded tab name using the Sheets collection
    Sheets("Ashlesh").Range("B1") = "Ashlesh"
    
    ' Write to a hardcoded tab name using the Worksheets collection
    Worksheets("Ashlesh").Range("a1:a10") = "Ashlesh"

End Sub