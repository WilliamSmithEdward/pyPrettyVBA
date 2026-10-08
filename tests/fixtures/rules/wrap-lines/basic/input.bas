Sub Demo()

    CopyRange Worksheets("Data").Range("A1"), Worksheets("Archive").Range("A1"), rowCount, columnCount
    message = "The report for " & regionName & " covers " & CStr(rowCount) & " rows."
    If startDate <= today And endDate >= today And regionName <> "All" Then Exit Sub
    total = total + LongFunctionOne(firstArgument, secondArgument) + LongFunctionTwo(third) ' running total
    url = "https://example.com/a/very/long/path/that/goes/on/without/any/break/point.html"
    SendReport recipients, _
               subject, body

End Sub
