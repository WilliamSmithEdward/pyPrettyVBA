Sub Spacing()
    Dim x As Long, s As String
    x = 1 + 2 * 3 - 4 / 5
    x = -x
    x = x - -1
    s = "a" & "b"
    If x <> 1 And x >= 0 Or Not (x = 5) Then x = 3
    For x = 10 To 1 Step -1
    Next
    Debug.Print x; s, x
    MsgBox prompt:="hi", buttons:=vbOKOnly
    Print #1, "a"; "b"
    Report x, s
    Report (x), s
    Report -1, s
    Report , s
    x = Compute(x)
    If x >= 2 Then x = 1
End Sub
