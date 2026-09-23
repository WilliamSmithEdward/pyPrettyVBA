Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Sub Demo()
    Debug.Print gettickcount
End Sub
