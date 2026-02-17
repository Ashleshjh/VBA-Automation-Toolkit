Option Explicit

Sub alignment()

Dim name As String
name = "Ashlesh"
Range("a1:a10") = name
Range("a1:a10").HorizontalAlignment = xlLeft
Range("a1:a10").HorizontalAlignment = xlRight
Range("a1:a10").HorizontalAlignment = xlCenter
Range("a1:a10").VerticalAlignment = xlTop
Range("a1:a10").VerticalAlignment = xlBottom
Range("a1:a10").VerticalAlignment = xlCenter

End Sub

