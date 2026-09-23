Option Explicit

Private mTotal As Long
Private Const MAX_ITEMS As Long = 100

Public Enum Colour
    ColourRed = 1
    ColourGreen
End Enum

Public Function AddItem(ByVal itemValue As Long) As Boolean
    If MTOTAL + ITEMVALUE > max_items Then
        addItem = False
        GoTo done
    End If
    mtotal = mtotal + itemvalue
    AddItem = (colourred <> COLOURGREEN)
Done:
    Worksheets(1).range("A1").value = mtotal
    MsgBox Prompt:="Added", title:="Items"
End Function
