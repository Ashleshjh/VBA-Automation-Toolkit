Option Explicit

Sub border()

Dim name As String
name = "Ashlesh"
Range("a1:a10") = name
With Range("a1:a10").Borders
    .LineStyle = xlDot
    .Color = vbGreen
    .Weight = 3
    .LineStyle = xlDash
    .LineStyle = xlContinuous
    .LineStyle = xlDouble
    .LineStyle = xlNone

End With

End Sub


