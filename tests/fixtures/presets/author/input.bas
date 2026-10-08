Option Explicit
Private count As Long, limit As Long
' Loads the sheet.
Public Sub Load(ws As Worksheet)
    On Error GoTo Handler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim r As Long, total As Double, message As String
    Let total = 0
    If ws Is Nothing Then
        Exit Sub
    End If
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row: total = total + ws.Cells(r, 2).Value: Next r
    If total > limit Then
        Call Warn(ws.Name, total)
    Else
        Debug.Print "ok"   'within limit
    End If
    message = "Loaded " & ws.Name & " with " & CStr(count) & " rows and a total of " & Format(total, "#,##0.00") & " for the period."
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
Handler:
    MsgBox Err.Description
End Sub
