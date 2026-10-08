Option Explicit
Private count As Long
Private Type Point
    X As Long
    Y As Long
End Type
' Resets the counter.
Public Sub Reset()
    count = 0
End Sub


Public Property Get Count() As Long
    Count = count
End Property
Public Sub Report(items As Collection)
    Dim item As Variant

    If items.Count = 0 Then Exit Sub
    For Each item In items
        If item > 10 Then
            Debug.Print "big"
        ElseIf item > 5 Then
            Debug.Print "medium"
        Else
            Debug.Print "small"
        End If
    Next item
    Select Case count
        Case 0
            Debug.Print "none"
        Case Else
            Debug.Print count
    End Select
    With Application
        .ScreenUpdating = True
    End With
End Sub
