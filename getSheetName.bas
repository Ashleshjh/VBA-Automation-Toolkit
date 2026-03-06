Option Explicit

Sub getsheetname()

    Dim i As Long
    Dim sheetname As String
    Dim wsOutput As Worksheet
    
    ' Assign the explicit target worksheet object
    Set wsOutput = ThisWorkbook.Sheets("Output")
    
    ' Retrieve and display hardcoded index names via MsgBox
    sheetname = Sheets(1).Name
    MsgBox (sheetname)
    
    sheetname = Sheets(5).Name
    MsgBox (sheetname)
    
    MsgBox (Sheets(3).Name)

    ' Loop through all sheets and safely write their names to the declared Output sheet
    For i = 1 To ThisWorkbook.Sheets.Count
        wsOutput.Cells(i + 1, 1).Value = ThisWorkbook.Sheets(i).Name
    Next i

    ' Display completion message and total count
    MsgBox "All sheet names listed in '" & wsOutput.Name & "'.", vbInformation
    MsgBox ThisWorkbook.Sheets.Count

End Sub