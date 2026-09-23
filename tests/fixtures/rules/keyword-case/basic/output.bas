Option Explicit
Option Compare Text
DefLng A-C, X-Z

Private Declare PtrSafe Function tickcount Lib "kernel32" Alias "GetTickCount" () As Long

Public Sub demo(ByVal count As Long)
    Dim i As Long
    For i = 1 To count Step 2
        If i Mod 3 = 0 Then
            Debug.Print "fizz"
        ElseIf i > 10 Then
            Exit For
        End If
    Next i
    On Error GoTo handler
    workbook.close savechanges:=False
    Exit Sub
handler:
    Rem a remark
    Resume Next
End Sub
