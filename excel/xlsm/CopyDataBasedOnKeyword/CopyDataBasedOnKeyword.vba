''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyDataBasedOnKeyword
''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyDataBasedOnKeyword()

    Dim wsInput As Worksheet
    Dim wsCopySource As Worksheet
    Dim lastRowInput As Long
    Dim lastRowCopySource As Long
    Dim i As Long
    Dim j As Long
    Dim targetKeyword As String
    Dim sourceKeyword As String
    Dim pasteRow As Long
    Dim keywordParts As Variant
    Dim keywordPart As Variant
    Dim insertCount As Long

    ' Set the source and destination worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsCopySource = ThisWorkbook.Sheets("CopySource")

    ' Clear all yellow-highlighted rows to ensure a clean start
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row
    For i = lastRowInput To 2 Step -1
        If wsInput.Cells(i, "A").Interior.Color = RGB(255, 255, 0) Then
            wsInput.Rows(i).Delete
        End If
    Next i

    ' Recalculate the last row after the cleanup
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "C").End(xlUp).Row
    lastRowCopySource = wsCopySource.Cells(wsCopySource.Rows.Count, "C").End(xlUp).Row

    ' Loop through each row in the Input sheet from bottom to top
    For i = lastRowInput To 2 Step -1
        ' Reset the insertion count for each 'sub' row
        insertCount = 0

        ' Check if the "Flg" column (G) value is "sub"
        If LCase(wsInput.Cells(i, "G").Value) = "sub" Then
            ' Get the keyword from the Input sheet
            targetKeyword = Trim(wsInput.Cells(i, "C").Value)

            ' Replace the newline character (vbLf) with a comma before splitting
            targetKeyword = Replace(targetKeyword, vbLf, ",")

            ' Handle cases with multiple keywords, now including multi-line cells
            keywordParts = Split(targetKeyword, "｣ｬ")
            If UBound(keywordParts) = 0 Then
                keywordParts = Split(targetKeyword, ",")
            End If

            ' Loop through each individual keyword part
            For Each keywordPart In keywordParts
                Dim cleanedTarget As String
                cleanedTarget = Replace(LCase(Trim(keywordPart)), " ", "")

                ' Loop through the CopySource sheet to find matches
                For j = 2 To lastRowCopySource
                    Dim cleanedSource As String
                    cleanedSource = Replace(LCase(Trim(wsCopySource.Cells(j, "C").Value)), " ", "")

                    If cleanedSource = cleanedTarget Then
                        ' Calculate the position for the new row insertion
                        pasteRow = i + 1 + insertCount

                        ' Insert a new row
                        wsInput.Rows(pasteRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        ' Copy the matching row
                        wsCopySource.Rows(j).Copy Destination:=wsInput.Rows(pasteRow)
                        ' Fill the new row with yellow color
                        wsInput.Rows(pasteRow).Interior.Color = RGB(255, 255, 0)

                        ' Increment the insertion count
                        insertCount = insertCount + 1
                    End If
                Next j
            Next keywordPart
        End If
    Next i

    ' Clear the clipboard
    Application.CutCopyMode = False
    ' Inform the user that the script is complete
    MsgBox "Data transfer and insertion complete!", vbInformation, "Done"

End Sub
