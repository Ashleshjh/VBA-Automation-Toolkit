Option Explicit

Sub AddBorders()
    ' Purpose: Applies various border styles to a range
    
    Dim name As String
    name = "Ashlesh"
    
    ' Assign value first
    Range("A1:A10").Value = name
    
    ' Use 'With' block to apply multiple properties to the same object efficiently
    With Range("A1:A10").Borders
        .LineStyle = xlDot         ' Dotted lines
        .Color = vbGreen           ' Green color
        .Weight = 3                ' Thicker border
        
        ' Other styles (Uncomment to test)
        ' .LineStyle = xlDash
        ' .LineStyle = xlContinuous
        ' .LineStyle = xlDouble
        
        ' Removes all borders
        ' .LineStyle = xlNone
    End With
End Sub
