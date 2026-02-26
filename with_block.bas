Option Explicit

Sub with_block()

    ' Declare a string variable (Warning: 'name' is a reserved VBA property)
    Dim name As String

    ' Assign the string value
    name = "Ashlesh"
    
    ' Implicitly assign the string to the default Value property of the range
    Range("A1:a10") = name

    ' Open a With block to hold the Font object in memory
    With Range("a1:a10").Font
        ' Set font type (Warning: Typo in font name)
        .name = "ariel"
        
        ' Scale size up to 20
        .Size = 20
        
        ' Apply formatting
        .Bold = True
        .Italic = True
        .Underline = True
        
        ' Immediately strip formatting
        .Bold = False
        .Italic = False
        .Underline = False
        
        ' Apply and immediately strip strikethrough
        .Strikethrough = True
        .Strikethrough = False
        
        ' Overwrite the previous size command and scale back down to 11
        .Size = 11
    End With

End Sub