Option Explicit

Sub sheetprotection()

    ' Lock the worksheet using a hardcoded integer as a password
    Sheets("cellreferencing").Protect Password:=123
    
    ' Immediately unlock the exact same worksheet in the next execution cycle
    Sheets("cellreferencing").Unprotect Password:=123

End Sub