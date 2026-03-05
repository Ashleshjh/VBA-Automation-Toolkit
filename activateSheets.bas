Option Explicit

Sub activatesheets()

    ' Activate sheet8
    Sheets("sheet8").Activate
    
    ' Instantly overwrite the active state by selecting a different sheet
    Sheets("renamed using sheets 4").Select

End Sub