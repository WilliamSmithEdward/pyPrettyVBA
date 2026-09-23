Sub Report()
    Dim s As String
    s = ucase$(trim$(environ$("USERNAME"))) & vbcrlf
    msgbox s, vbokonly + vbinformation, "Report"
    Debug.Print format$(now, "yyyy-mm-dd"), Len(s)
    err.raise vbobjecterror + 1, "Report", err.description
    activesheet.Range("A1").Value = s
#If vba7 And win64 Then
    Debug.Print "64-bit"
#End If
End Sub
