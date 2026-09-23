Sub Tidy()
    Dim doc As document
    Set doc = activedocument
    selection.EndKey Unit:=wdline
    doc.Range.InsertAfter "done"
End Sub
