Public Sub Load()
    On Error GoTo Handler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    DoWork
    On Error Resume Next
    Set ws = Worksheets("Data")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    On Error GoTo Handler
    Application.Run "Rebuild"
    If big Then
        Application.Calculation = xlCalculationManual
        DoWork
    End If
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub
