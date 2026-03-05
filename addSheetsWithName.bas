Option Explicit

Sub addsheetswithname()

    ' Single-line creation and naming method (commented out)
    'Worksheets.Add.Name = "addedwithVBA"
    
    ' Declare a Worksheet object variable
    Dim ws As Worksheet
    
    ' Instantiate the new worksheet and capture it in memory
    Set ws = Worksheets.Add
    
    ' Assign a specific string name to the captured worksheet object
    ws.Name = "addedwithVBA"

End Sub