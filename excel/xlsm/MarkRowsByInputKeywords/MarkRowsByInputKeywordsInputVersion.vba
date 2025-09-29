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

Sub TrimKeywords(ByRef keywords() As String)
' Function: Removes leading and trailing spaces from every element in the keyword array.

    Dim i As Long

    For i = LBound(keywords) To UBound(keywords)
        keywords(i) = Trim(keywords(i))
    Next i

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''
' MarkRowsByInputKeywords(main)
''''''''''''''''''''''''''''''''''''''''''''''''''

Sub MarkRowsByInputKeywords()
    Dim ws As Worksheet
    Dim bounds As Variant
    Dim lastRow As Long, lastCol As Long, rowNum As Long, colNum As Long, markCol As Long
    Dim inputStr As String
    Dim keywords() As String
    Dim keyword As Variant ' Loop variable for keywords
    Dim found As Boolean

    ' 1. Input Keywords
    inputStr = InputBox("Enter keywords separated by commas (e.g.: apple, banana, orange)", "Keyword Input")

    ' Check for empty input
    If Trim(inputStr) = "" Then
        MsgBox "No keywords entered. Operation cancelled.", vbInformation
        Exit Sub
    End If

    ' 2. Split and Clean Keywords
    keywords = Split(inputStr, ",")
    Call TrimKeywords(keywords) ' Clean up spaces from keywords

    ' 3. Set Worksheet and Determine Bounds
    Set ws = ThisWorkbook.Sheets("MarkRowsByKeywords") ' **Modify "Sheet1" to your actual sheet name**

    ' Dynamically determine the last row and column
    bounds = GetSheetBounds(ws)
    lastRow = bounds(0)
    lastCol = bounds(1)
    Debug.Print "Worksheet: " & ws.Name & _
            ", MaxRow=" & lastRow & _
            ", MaxColumn=" & lastCol
    ' Set the column for marking (Z is column 26)
    markCol = 1 + lastCol

    ' 4. Loop Through Rows and Columns
    ' Start from row 2 (assuming row 1 is the header)
    For rowNum = 2 To lastRow
        found = False

        ' Iterate through all data columns (from A column/1 to lastCol)
        For colNum = 1 To lastCol

            ' Check if the cell has a value
            If ws.Cells(rowNum, colNum).Value <> "" Then

                ' Iterate through all cleaned keywords
                For Each keyword In keywords
                    ' Use InStr for case-insensitive substring matching
                    If InStr(1, ws.Cells(rowNum, colNum).Value, keyword, vbTextCompare) > 0 Then

                        ' Match found: Mark the row and exit inner loops
                        ws.Cells(rowNum, markCol).Value = "Matched"
                        found = True
                        Exit For ' Exit keyword loop
                    End If
                Next keyword
            End If

            If found Then Exit For ' Exit column loop, move to the next row
        Next colNum
    Next rowNum

    MsgBox "Processing complete! Rows containing keywords have been marked in column " & Split(Cells(1, markCol).Address, "$")(1) & ".", vbInformation
End Sub
