Option Explicit

Private Sub Worksheet_Activate()
    ' Populate student list from DB sheet
    Call PopulateStudents
    Call RefreshListBox
End Sub

Private Sub PopulateStudents()
    Dim wsDB As Worksheet, lastRow As Long, i As Long
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set wsDB = ThisWorkbook.Sheets("DB")
    
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    cmbStudent.clear
    
    For i = 2 To WorksheetFunction.Max(2, lastRow)
        If Trim(wsDB.Cells(i, "A").Value) <> "" Then
            If Not dict.Exists(wsDB.Cells(i, "A").Value) Then
                dict.Add wsDB.Cells(i, "A").Value, 1
                cmbStudent.AddItem wsDB.Cells(i, "A").Value
            End If
        End If
    Next i
End Sub

Private Sub scrMarks_Change()
    txtMarks.Value = scrMarks.Value
End Sub

Private Sub btnAdd_Click()
    Dim wsDB As Worksheet, lastRow As Long, stName As String
    Dim marks As Long, result As String
    Dim passMark As Integer
    
    Set wsDB = ThisWorkbook.Sheets("DB")
    stName = Trim(Me.cmbStudent.Value)
    
    If stName = "" Then
        MsgBox "Enter or select a student name.", vbExclamation
        Exit Sub
    End If
    
    If Not IsNumeric(Me.txtMarks.Value) Then
        MsgBox "Enter numeric marks (0-100).", vbExclamation
        Exit Sub
    End If
    
    marks = CLng(Me.txtMarks.Value)
    If marks < 0 Or marks > 100 Then
        MsgBox "Marks must be between 0 and 100.", vbExclamation
        Exit Sub
    End If
    
    ' Passing scale based on strict/lenient
    If Me.optStrict.Value = True Then
        passMark = 35
    ElseIf Me.optLenient.Value = True Then
        passMark = 30
    Else
        MsgBox "Select grading scale (Strict or Lenient).", vbExclamation
        Exit Sub
    End If
    
    ' Determine result
    If marks >= passMark Then
        result = "Pass"
    Else
        result = "Fail"
    End If
    
    ' Append to DB sheet
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
    wsDB.Cells(lastRow, "A").Value = stName
    wsDB.Cells(lastRow, "B").Value = marks
    wsDB.Cells(lastRow, "C").Value = result
    
    MsgBox "Saved: " & stName & " - " & marks & " (" & result & ")", vbInformation, "Success"
    
    ' Optional: clear fields
    Me.cmbStudent.Value = ""
    Me.txtMarks.Value = ""
    Me.optStrict.Value = False
    Me.optLenient.Value = False
End Sub

Private Sub RefreshListBox()
    Dim wsDB As Worksheet, lastRow As Long, i As Long
    Set wsDB = ThisWorkbook.Sheets("DB")
    
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    lstRecords.clear
    
    For i = 2 To WorksheetFunction.Max(2, lastRow)
        If Trim(wsDB.Cells(i, "A").Value) <> "" Then
            lstRecords.AddItem wsDB.Cells(i, "A").Value & " | " & wsDB.Cells(i, "B").Value & " | " & wsDB.Cells(i, "C").Value
        End If
    Next i
    
    ' Repopulate ComboBox (keeps list updated)
    Call PopulateStudents
End Sub



