Option Explicit

Sub forand()
    ' Declare integer variable for the loop counter
    Dim i As Integer
    
    ' Loop through rows 2 to 11
    For i = 2 To 11
        
        ' Check if the value in column A is between 1 and 34
        If Cells(i, 1).Value >= 1 And Cells(i, 1).Value <= 34 Then
            ' Write "Fail" in column B if true
            Cells(i, 2).Value = "Fail"
            
        ' Check if the value in column A is between 35 and 100
        ElseIf Cells(i, 1).Value >= 35 And Cells(i, 1).Value <= 100 Then
            ' Write "Pass" in column B if true
            Cells(i, 2).Value = "Pass"
            
        ' Catch any other values (e.g., negative numbers, over 100, or text)
        Else
            ' Write "Invalid" in column B for out-of-range values
            Cells(i, 2).Value = "Invalid"
            
        End If
        
    Next i
End Sub