Private Declare PtrSafe Function CopyBytes Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByVal dest As LongPtr, _
    ByVal src As LongPtr _
) As Long

Sub Query()
    Dim sql As String
    sql = "SELECT * " & _
          "FROM t " & _
          "WHERE x = 1"
    Report sql, _
        1, _
        2
End Sub
