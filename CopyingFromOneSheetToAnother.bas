Sub CopyingFromOneSheetToAnother()
    ' Loop through rows 2 to 5
    For i = 2 To 5
        
        ' Copy column A from "dept" sheet to column A in "dept-emp" sheet
        Worksheets("dept-emp").Cells(i, "A") = Worksheets("dept").Cells(i, "A")
        
        ' Copy column A from "emp" sheet to column B in "dept-emp" sheet
        Worksheets("dept-emp").Cells(i, "B") = Worksheets("emp").Cells(i, "A")
        
    Next i
End Sub