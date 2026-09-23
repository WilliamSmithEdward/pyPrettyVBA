Sub Demo()
    x = 1
    y = 2
    z = 3
    If ready Then total = total + 1: count = 0
    For i = 1 To 3
        s = s + i
    Next i
Retry: attempts = attempts + 1
    Select Case x
        Case 1
            y = 10
    End Select
End Sub
