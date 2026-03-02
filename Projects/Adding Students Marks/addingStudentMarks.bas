Option Explicit

Private Sub Worksheet_Activate()

    ' Triggered when the worksheet is opened or switched to
    ' Populate student list from DB sheet
    Call PopulateStudents
    
    ' Update the displayed records in the list box
    Call RefreshListBox

End Sub

Private Sub PopulateStudents()

    Dim wsDB As Worksheet, lastRow As Long, i As Long
    Dim dict As Object
    
    ' Create a dictionary to store unique student names
    Set dict = CreateObject("Scripting.Dictionary")
    Set wsDB = ThisWorkbook.Sheets("DB")
    
    ' Find the last used row in Column A of the DB sheet
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    ' Clear existing items in the ComboBox
    cmbStudent.clear
    
    ' Loop through the DB sheet to extract unique student names
    For i = 2 To WorksheetFunction.Max(2, lastRow)
        If Trim(wsDB.Cells(i, "A").Value) <> "" Then
            ' If the name is not in the dictionary, add it and populate the ComboBox
            If Not dict.Exists(wsDB.Cells(i, "A").Value) Then
                dict.Add wsDB.Cells(i, "A").Value, 1
                cmbStudent.AddItem wsDB.Cells(i, "A").Value
            End If
        End If
    Next i

End Sub

Private Sub scrMarks_Change()

    ' Link the scrollbar value directly to the textbox display
    txtMarks.Value = scrMarks.Value

End Sub

Private Sub btnAdd_Click()

    Dim wsDB As Worksheet, lastRow As Long, stName As String
    Dim marks As Long, result As String
    Dim passMark As Integer
    
    Set wsDB = ThisWorkbook.Sheets("DB")
    stName = Trim(Me.cmbStudent.Value)
    
    ' Validate that a student name has been selected or entered
    If stName = "" Then
        MsgBox "Enter or select a student name.", vbExclamation
        Exit Sub
    End If
    
    ' Validate that the entered marks are numeric
    If Not IsNumeric(Me.txtMarks.Value) Then
        MsgBox "Enter numeric marks (0-100).", vbExclamation
        Exit Sub
    End If
    
    ' Convert text to Long and validate the 0-100 range bounds
    marks = CLng(Me.txtMarks.Value)
    If marks < 0 Or marks > 100 Then
        MsgBox "Marks must be between 0 and 100.", vbExclamation
        Exit Sub
    End If
    
    ' Determine the passing threshold based on selected Option Button
    If Me.optStrict.Value = True Then
        passMark = 35
    ElseIf Me.optLenient.Value = True Then
        passMark = 30
    Else
        MsgBox "Select grading scale (Strict or Lenient).", vbExclamation
        Exit Sub
    End If
    
    ' Evaluate the student's marks against the selected threshold
    If marks >= passMark Then
        result = "Pass"
    Else
        result = "Fail"
    End If
    
    ' Find the next empty row in the DB sheet and append the new record
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row + 1
    wsDB.Cells(lastRow, "A").Value = stName
    wsDB.Cells(lastRow, "B").Value = marks
    wsDB.Cells(lastRow, "C").Value = result
    
    ' Confirm successful entry to the user
    MsgBox "Saved: " & stName & " - " & marks & " (" & result & ")", vbInformation, "Success"
    
    ' Reset the form fields for the next entry
    Me.cmbStudent.Value = ""
    Me.txtMarks.Value = ""
    Me.optStrict.Value = False
    Me.optLenient.Value = False

End Sub

Private Sub RefreshListBox()

    Dim wsDB As Worksheet, lastRow As Long, i As Long
    Set wsDB = ThisWorkbook.Sheets("DB")
    
    lastRow = wsDB.Cells(wsDB.Rows.Count, "A").End(xlUp).Row
    
    ' Clear the current ListBox contents
    lstRecords.clear
    
    ' Loop through the DB sheet and concatenate columns into single ListBox strings
    For i = 2 To WorksheetFunction.Max(2, lastRow)
        If Trim(wsDB.Cells(i, "A").Value) <> "" Then
            lstRecords.AddItem wsDB.Cells(i, "A").Value & " | " & wsDB.Cells(i, "B").Value & " | " & wsDB.Cells(i, "C").Value
        End If
    Next i
    
    ' Repopulate ComboBox (keeps list updated)
    Call PopulateStudents

End Sub
