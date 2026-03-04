Option Explicit

Private Sub cmdSubmit_Click()

    Dim username As String
    Dim gender As String
    Dim skills As String

    ' Extract and trim the name input
    username = Trim(txtName.Value)

    ' Validate that the user entered a name
    If username = "" Then
        MsgBox "Please enter your name.", vbExclamation, "Missing Info"
        Exit Sub
    End If

    ' Determine the selected gender
    If optMale.Value = True Then
        gender = "Male"
    ElseIf optFemale.Value = True Then
        gender = "Female"
    Else
        gender = "Not specified"
    End If

    ' Concatenate selected skills with a trailing comma and space
    If chkExcel.Value = True Then skills = skills & "Excel, "
    If chkVBA.Value = True Then skills = skills & "VBA, "
    If chkSQL.Value = True Then skills = skills & "SQL, "

    ' Strip the trailing comma and space if skills were selected, otherwise default to "None"
    If skills <> "" Then 
        skills = Left(skills, Len(skills) - 2) 
    Else 
        skills = "None"
    End If

    ' Display the final formatted summary to the user
    MsgBox "Hi " & username & "! You selected " & gender & " and your skills are: " & skills, vbInformation, "Summary"

End Sub
