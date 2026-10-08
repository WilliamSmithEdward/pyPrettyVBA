Sub Demo()

    If x Then

        Exit Sub

    End If

    If y > 0 Then
        total = total + y
    Else
        total = 0
    End If

    If z Then
        ' explain
        Reset
    End If

    If ready Then
        DoVeryLongProcedureNameNumberOne SomeArgumentNameThatIsLong, AnotherArgumentNameThatIsLong, YetAnotherArgument
    End If

End Sub
