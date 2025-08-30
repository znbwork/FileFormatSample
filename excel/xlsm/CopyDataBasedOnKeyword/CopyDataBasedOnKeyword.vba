''''''''''''''''''''''''''''''''''''''''''''''''''
' CopyDataBasedOnKeyword
''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CopyDataBasedOnKeyword()

    ' Declare variables
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
    Dim pasteOffset As Long

    ' Set the source and destination worksheets
    Set wsInput = ThisWorkbook.Sheets("Input")
    Set wsCopySource = ThisWorkbook.Sheets("CopySource")

    ' --- NEW: Clear all yellow-highlighted rows first ---
    ' Find the last row to check for yellow cells
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "A").End(xlUp).Row

    ' Loop from the bottom up to delete any yellow-highlighted rows
    For i = lastRowInput To 2 Step -1
        If wsInput.Cells(i, "A").Interior.Color = RGB(255, 255, 0) Then
            wsInput.Rows(i).Delete
        End If
    Next i

    ' --- End of the cleanup section ---

    ' Recalculate the last row after deleting the old data
    lastRowInput = wsInput.Cells(wsInput.Rows.Count, "C").End(xlUp).Row
    lastRowCopySource = wsCopySource.Cells(wsCopySource.Rows.Count, "C").End(xlUp).Row

    ' Loop through each row in the Input sheet from bottom to top
    For i = lastRowInput To 2 Step -1
        ' Check if the "Flg" column (G) value is "sub"
        If LCase(wsInput.Cells(i, "G").Value) = "sub" Then
            ' Get the keyword from the Input sheet
            targetKeyword = Trim(wsInput.Cells(i, "C").Value)

            ' Handle multiple keywords separated by commas
            ' First, try splitting by Chinese comma
            keywordParts = Split(targetKeyword, "，")
            ' If that fails, try splitting by English comma
            If UBound(keywordParts) = 0 Then
                keywordParts = Split(targetKeyword, ",")
            End If

            pasteOffset = 0

            ' Loop through each individual keyword
            For Each keywordPart In keywordParts
                keywordPart = Trim(keywordPart)

                ' Loop through each row in the CopySource sheet
                For j = 2 To lastRowCopySource
                    sourceKeyword = Trim(wsCopySource.Cells(j, "C").Value)

                    ' Check if the keyword from CopySource matches the target keyword
                    If LCase(sourceKeyword) = LCase(keywordPart) Then
                        ' The paste row is the current row (i) plus the accumulated offset
                        pasteRow = i + 1 + pasteOffset

                        ' Insert a new row below the current row
                        wsInput.Rows(pasteRow).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        ' Copy the entire row from CopySource to the newly inserted row
                        wsCopySource.Rows(j).Copy Destination:=wsInput.Rows(pasteRow)
                        ' Apply yellow fill color to the pasted row
                        wsInput.Rows(pasteRow).Interior.Color = RGB(255, 255, 0)

                        ' Increment the paste offset for subsequent insertions at this spot
                        pasteOffset = pasteOffset + 1
                    End If
                Next j
            Next keywordPart
        End If
    Next i

    ' Clear the clipboard
    Application.CutCopyMode = False
    ' Optional: Provide a message to the user that the script is complete
    MsgBox "Data transfer and insertion complete!", vbInformation, "Done"

End Sub
