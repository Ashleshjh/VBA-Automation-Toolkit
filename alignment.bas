Option Explicit

Sub ChangeAlignment()
    ' Purpose: Adjusts text alignment (Horizontal and Vertical)
    
    Dim name As String
    name = "Ashlesh"
    
    Range("A1:A10").Value = name
    
    ' Horizontal Alignment Options
    Range("A1:A10").HorizontalAlignment = xlLeft
    Range("A1:A10").HorizontalAlignment = xlRight
    Range("A1:A10").HorizontalAlignment = xlCenter
    
    ' Vertical Alignment Options
    Range("A1:A10").VerticalAlignment = xlTop
    Range("A1:A10").VerticalAlignment = xlBottom
    Range("A1:A10").VerticalAlignment = xlCenter
End Sub
