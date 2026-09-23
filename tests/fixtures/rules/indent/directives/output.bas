#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)
#End If

#If VBA7 Then
Public Sub Pause(ByVal ms As LongPtr)
#Else
Public Sub Pause(ByVal ms As Long)
#End If
    If ms > 0 Then
        #If DEBUG_MODE Then
            Debug.Print "pausing"; ms
        #End If
        Sleep ms
    End If
End Sub
