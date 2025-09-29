Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''
' GetSheetBounds
''''''''''''''''''''''''''''''''''''''''''''''''''

Function GetSheetBounds(ByVal ws As Worksheet) As Long()
    Dim bounds(1) As Long
    Dim r As Range

    On Error GoTo EmptySheet
    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not r Is Nothing Then
        bounds(0) = r.Row
    Else
        GoTo EmptySheet
    End If

    Set r = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not r Is Nothing Then
        bounds(1) = r.Column
    Else
        GoTo EmptySheet
    End If

    GetSheetBounds = bounds
    Exit Function

EmptySheet:
    bounds(0) = 1
    bounds(1) = 1
    GetSheetBounds = bounds
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''
' TrimKeywords
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub TrimKeywords(ByRef keywords As Variant)
    Dim i As Long
    If Not IsArray(keywords) Then Exit Sub
    On Error Resume Next
    For i = LBound(keywords) To UBound(keywords)
        keywords(i) = Trim(CStr(keywords(i)))
    Next i
    On Error GoTo 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''
' MarkRowsByInputKeywords(main)
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub MarkRowsByInputKeywords()
    Dim ws As Worksheet
    Dim bounds As Variant
    Dim lastRow As Long, lastCol As Long, rowNum As Long, colNum As Long, markCol As Long
    Dim keywords As Variant
    Dim keyword As Variant
    Dim found As Boolean

    ' Fixed keywords (no input box)
    keywords = Array("apple", "banana", "orange")
    Call TrimKeywords(keywords)

    ' Set worksheet - change name if needed
    Set ws = ThisWorkbook.Sheets("MarkRowsByKeywords")

    ' Determine bounds
    bounds = GetSheetBounds(ws)
    lastRow = bounds(0)
    lastCol = bounds(1)

    Debug.Print "Worksheet: " & ws.Name & ", MaxRow=" & lastRow & ", MaxColumn=" & lastCol

    If lastRow < 2 Then
        MsgBox "No data rows found (lastRow=" & lastRow & ").", vbInformation, "Info"
        Exit Sub
    End If

    ' Mark column is the next column after lastCol
    markCol = lastCol + 1

    For rowNum = 2 To lastRow
        found = False
        For colNum = 1 To lastCol
            If Len(ws.Cells(rowNum, colNum).Value & "") > 0 Then
                For Each keyword In keywords
                    If InStr(1, ws.Cells(rowNum, colNum).Value, CStr(keyword), vbTextCompare) > 0 Then
                        ws.Cells(rowNum, markCol).Value = "Matched"
                        found = True
                        Exit For
                    End If
                Next keyword
            End If
            If found Then Exit For
        Next colNum
    Next rowNum

    MsgBox "Processing complete! Rows containing keywords have been marked in column " & Split(ws.Cells(1, markCol).Address, "$")(1) & ".", vbInformation, "Done"
End Sub
