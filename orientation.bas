Sub orientation()

    ' Set text to horizontal (default)
    Range("a1").orientation = 0
    
    ' Tilt text upwards in 10-degree increments
    Range("a1").orientation = 10
    Range("a1").orientation = 20
    Range("a1").orientation = 30
    Range("a1").orientation = 40
    Range("a1").orientation = 50
    Range("a1").orientation = 60
    Range("a1").orientation = 70
    Range("a1").orientation = 80
    
    ' Set text completely vertical (facing up)
    Range("a1").orientation = 90
    
    ' Tilt text downwards in 10-degree increments
    Range("a1").orientation = -10
    Range("a1").orientation = -20
    Range("a1").orientation = -30
    Range("a1").orientation = -40
    Range("a1").orientation = -50
    Range("a1").orientation = -60
    Range("a1").orientation = -70
    Range("a1").orientation = -80
    
    ' Set text completely vertical (facing down)
    Range("a1").orientation = -90

End Sub