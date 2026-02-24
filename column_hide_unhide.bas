Option Explicit

Sub column_hide_unhide()

    ' Hide Column A
    Range("a:a").Columns.Hidden = True
    
    ' Unhide Column A
    Range("a:a").Columns.Hidden = False
    
    ' Hide Columns A through M
    Range("a:m").Columns.Hidden = True
    
    ' Unhide Columns A through M
    Range("a:m").Columns.Hidden = False

End Sub