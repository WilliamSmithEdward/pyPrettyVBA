Option Explicit
Private Type Point
	x As Long
	y As Long
End Type
Private Enum Direction
	North
	South
End Enum
Public Function Classify(ByVal n As Long) As String
	Dim i As Long
	If n < 0 Then
		Classify = "negative"
	ElseIf n = 0 Then
		Classify = "zero"
	Else
		Select Case n
			Case 1 To 9
				Classify = "small"
			Case Else
				For i = 1 To n
					Do While i < 3
						i = i + 1
					Loop
				Next i
				With Application
					.StatusBar = "busy"
				End With
				Classify = "large"
		End Select
	End If
	If n = 42 Then Classify = "answer"
End Function
