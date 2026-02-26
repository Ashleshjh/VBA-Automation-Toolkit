Option Explicit

Sub wraptext()

    ' Enable text wrapping for cells A1 through A10
    Range("a1:a10").wraptext = True
    
    ' Immediately disable text wrapping, overwriting the previous command
    Range("a1:a10").wraptext = False

End Sub