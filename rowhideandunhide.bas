Option Explicit

Sub rowhideandunhide()

    ' Hide Row 1
    Range("1:1").Rows.Hidden = True
    
    ' Unhide Row 1
    Range("1:1").Rows.Hidden = False
    
    ' Hide Rows 1 through 5
    Range("1:5").Rows.Hidden = True
    
    ' Unhide Rows 1 through 5
    Range("1:5").Rows.Hidden = False

End Sub