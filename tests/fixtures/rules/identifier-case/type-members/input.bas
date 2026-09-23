Private Type Span
    start As Long
    second As Long
End Type

Sub Stamp()
    Dim s As Span
    s.second = SECOND(Now)
    s.start = 1
End Sub
