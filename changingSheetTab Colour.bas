Option Explicit

Sub changesheettabcolor()

    ' Sequentially overwrite the tab color 8 times using VBA color constants
    Sheets("sheet8").Tab.Color = vbWhite
    Sheets("sheet8").Tab.Color = vbBlack
    Sheets("sheet8").Tab.Color = vbRed
    Sheets("sheet8").Tab.Color = vbGreen
    Sheets("sheet8").Tab.Color = vbBlue
    Sheets("sheet8").Tab.Color = vbYellow
    Sheets("sheet8").Tab.Color = vbCyan
    Sheets("sheet8").Tab.Color = vbMagenta
    
    ' Immediately overwrite the tab color 3 more times using the legacy ColorIndex palette
    Sheets("sheet8").Tab.ColorIndex = 1
    Sheets("sheet8").Tab.ColorIndex = 50
    Sheets("sheet8").Tab.ColorIndex = 56
    
    ' Improperly assign a Boolean value to a Long data type property
    Sheets("sheet8").Tab.Color = False

End Sub