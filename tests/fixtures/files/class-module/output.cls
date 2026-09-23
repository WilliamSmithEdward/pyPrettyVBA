VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mValue As Long

Public Property Get value() As Long
Attribute value.VB_UserMemId = 0
    value = mValue
End Property

Public Sub increment()
    mValue = mValue + 1
End Sub
