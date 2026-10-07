Option Explicit

Private count As Long
Private limit As Long
Public WithEvents app As Application
Public name As String

Sub Demo()
    Dim ws As Worksheet
    Dim r As Long
    Dim total As Double
    Dim buffer(1 To 10) As Byte
    Dim flag As Boolean  ' scratch
    Dim x
    Dim y As Long
    Dim ok As New Collection
    Dim cell As Range
    Const A = 1
    Const B As String = "b"
    Static hits As Long
    Static misses As Long
    ReDim items(0 To r), names(0 To r)
    Dim one As Long
End Sub
