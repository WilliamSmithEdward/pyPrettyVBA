Sub Report()
    Dim s As String
    s = UCase$(Trim$(Environ$("USERNAME"))) & vbCrLf
    MsgBox s, vbOKOnly + vbInformation, "Report"
    Debug.Print Format$(Now, "yyyy-mm-dd"), Len(s)
    Err.Raise vbObjectError + 1, "Report", Err.Description
    ActiveSheet.Range("A1").Value = s
#If VBA7 And Win64 Then
    Debug.Print "64-bit"
#End If
End Sub
