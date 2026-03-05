Option Explicit

Sub addsheets()

    ' Add a new sheet immediately before the specified hardcoded tab
    Sheets.Add before:=Worksheets("Ashlesh-cellreferencing")
    
    ' Add a new worksheet immediately after the specified hardcoded tab
    Worksheets.Add after:=Worksheets("Ashlesh-cellreferencing")
    
    ' Add a new worksheet at the absolute end of the workbook
    Worksheets.Add after:=Worksheets(Worksheets.Count)

End Sub