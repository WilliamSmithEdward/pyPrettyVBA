Option Explicit

Private count As Long, limit As Long
Public WithEvents app As Application, name As String

Sub Demo()
    Dim ws As Worksheet, r As Long, total As Double
    Dim buffer(1 To 10) As Byte, flag As Boolean  ' scratch
    Dim x, y As Long
    Dim ok As New Collection, cell As Range
    Const A = 1, B As String = "b"
    Static hits As Long, misses As Long
    ReDim items(0 To r), names(0 To r)
    Dim one As Long
End Sub
