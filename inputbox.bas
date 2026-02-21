Option Explicit

Sub inputBox()
    ' Declare a string variable to store the text input
    Dim name As String
    
    ' Prompt the user with a dialog box and store their answer in the 'name' variable
    name = inputBox("Enter a name: ")
    
    ' Display a pop-up message combining static text with the entered variable
    MsgBox "Hi " & name & ", how are you?"
End Sub