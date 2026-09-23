Sub Export()
    Dim ws As Worksheet
    Set ws = activeworkbook.worksheets("Data")
    sql = "SELECT *"
    ws.cells(1, 1).value = sql
    ws.range("A1").end(xlup).select
    activeworkbook.close savechanges:=false
    msgbox "Done", title:="Export"
End Sub
