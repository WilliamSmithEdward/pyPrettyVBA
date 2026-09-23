Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
Private Declare PtrSafe Sub sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)

Public Sub Wait(ByVal ms As Long)
    Dim started As Long
    started = gettickcount
    Do While GETTICKCOUNT - started < ms
        SLEEP 10
    Loop
End Sub
