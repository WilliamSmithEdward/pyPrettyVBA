Sub Demo()
    On Error GoTo Fail
10  x = 1
20  y = 2
    If x Then
Retry:  x = x + 1
    End If
    Exit Sub
Fail:
    Resume Next
End Sub
