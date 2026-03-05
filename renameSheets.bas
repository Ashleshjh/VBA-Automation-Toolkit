Option Explicit

Sub renameasheet()

    ' Rename a worksheet by targeting its hardcoded string name
    Sheets("addedwithVBA").Name = "addedwithVBArenamed"
    
    ' Rename a worksheet by targeting its absolute index position
    Sheets(4).Name = "renamed using sheets 4"

End Sub