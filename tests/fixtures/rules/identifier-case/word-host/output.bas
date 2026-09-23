Sub Tidy()
    Dim doc As Document
    Set doc = ActiveDocument
    Selection.EndKey Unit:=wdLine
    doc.Range.InsertAfter "done"
End Sub
