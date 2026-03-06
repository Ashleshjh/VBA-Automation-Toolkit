Option Explicit

Sub hideunhidesheets()

    ' Instantly make the hardcoded sheet visible
    Sheets("sheet8").Visible = True
    
    ' Immediately hide it in the exact next millisecond
    Sheets("sheet8").Visible = False

End Sub