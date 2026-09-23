VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Greeter
   Caption         =   "Greeter"
   ClientHeight    =   3015
   ClientWidth     =   4560
   OleObjectBlob   =   "Greeter.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Greeter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGreet_Click()
    MsgBox "Hello, " & txtName.Text
End Sub
