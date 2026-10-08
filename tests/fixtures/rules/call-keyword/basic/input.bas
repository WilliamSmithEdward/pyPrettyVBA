Sub Demo()

    Call Reset
    Call Warn(ws.Name, total)
    Call ws.Range("A1").Select
    Call Log((message))
    Call Save()

End Sub
