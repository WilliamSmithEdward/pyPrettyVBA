Sub Export()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Worksheets("Data")
    Sql = "SELECT *"
    ws.Cells(1, 1).Value = Sql
    ws.Range("A1").End(xlUp).Select
    ActiveWorkbook.Close savechanges:=false
    MsgBox "Done", Title:="Export"
End Sub
