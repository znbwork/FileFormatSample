''''''''''''''''''''''''''''''''''''''''''''''''''
' MarkRowsByKeywords
''''''''''''''''''''''''''''''''''''''''''''''''''
Sub MarkRowsByKeywords()
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim keywords As Variant
    Dim cell As Range, k As Variant
    Dim found As Boolean

    ' Modify here: Keyword list (feel free to add or remove)
    keywords = Array("apple", "banana", "orange")

    ' Modify here: Target worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1")

    ' Find the last row (based on column A, can be changed to another column)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Iterate through each row
    For i = 2 To lastRow   ' Start from the second row (the first row is usually the header)
        found = False
        ' Iterate through columns A to Y of the current row
        For Each cell In ws.Range("A" & i & ":Y" & i)
            If cell.Value <> "" Then
                For Each k In keywords
                    If InStr(1, cell.Value, k, vbTextCompare) > 0 Then
                        ws.Cells(i, "Z").Value = "m"   ' Mark in column Z
                        found = True
                        Exit For
                    End If
                Next k
            End If
            If found Then Exit For
        Next cell
    Next i

    MsgBox "Processing complete! Rows containing keywords have been marked in column Z."
End Sub