Option Explicit

Private mTotal As Long
Private Const MAX_ITEMS As Long = 100

Public Enum Colour
    ColourRed = 1
    ColourGreen
End Enum

Public Function AddItem(ByVal itemValue As Long) As Boolean
    If mTotal + itemValue > MAX_ITEMS Then
        AddItem = False
        GoTo Done
    End If
    mTotal = mTotal + itemValue
    AddItem = (ColourRed <> ColourGreen)
Done:
    Worksheets(1).Range("A1").Value = mTotal
    MsgBox Prompt:="Added", Title:="Items"
End Function
