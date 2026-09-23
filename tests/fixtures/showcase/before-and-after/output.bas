Attribute VB_Name = "Inventory"
Option Explicit
'  Stock levels for the warehouse sheet.
Private Const LOW_STOCK = 10
Private mItems As Collection

Public Function RestockList(ws As Worksheet) As Collection
    Dim r As Long, qty As Long, result As New Collection
    On Error GoTo fail
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        qty = ws.Cells(r, 3).Value
        If qty < LOW_STOCK Then
            result.Add ws.Cells(r, 1).Value
        ElseIf qty = 0 Then result.Add ws.Cells(r, 1).Value & " (out)"
        End If
    Next r
    Set RestockList = result
    Exit Function
fail:
    MsgBox "Could not read row " & r & ": " & Err.Description, vbExclamation
End Function
