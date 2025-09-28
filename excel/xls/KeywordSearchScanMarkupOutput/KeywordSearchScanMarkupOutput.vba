Option Explicit

Sub KeywordSearchScanMarkupOutput()
    Dim wsMaster As Worksheet
    Dim ws As Worksheet
    Dim keywords As Collection
    Dim kw As Variant
    Dim kRow As Long

    Dim tmpVal As Variant
    Dim maxColParam As Long, maxRowParam As Long
    Dim workbookMaxCols As Long, workbookMaxRows As Long

    Dim usedFirstRow As Long, usedLastRow As Long
    Dim usedFirstCol As Long, usedLastCol As Long
    Dim scanLastRow As Long, scanLastCol As Long

    Dim r As Long, c As Long
    Dim val As String
    Dim outputFilePath As String
    Dim fileNum As Integer
    Dim outputText As String
    Dim matchFound As Boolean

    On Error GoTo ErrHandler

    ' master sheet
    On Error Resume Next
    Set wsMaster = ThisWorkbook.Worksheets("Master")
    On Error GoTo ErrHandler
    If wsMaster Is Nothing Then
        MsgBox "Sheet 'Master' not found!", vbCritical
        Exit Sub
    End If

    ' read keywords from B2 downward
    Set keywords = New Collection
    kRow = 2
    Do While Trim(CStr(wsMaster.Cells(kRow, 2).Value)) <> ""
        keywords.Add Trim(CStr(wsMaster.Cells(kRow, 2).Value))
        kRow = kRow + 1
    Loop
    If keywords.Count = 0 Then
        MsgBox "No keywords found in Master!B2:B", vbCritical
        Exit Sub
    End If

    ' read parameters: C2 = max column, D2 = max row
    tmpVal = wsMaster.Cells(2, 3).Value ' C2
    If IsNumeric(tmpVal) And tmpVal > 0 Then
        maxColParam = CLng(tmpVal)
    Else
        maxColParam = 0 ' auto detect per sheet
    End If

    tmpVal = wsMaster.Cells(2, 4).Value ' D2
    If IsNumeric(tmpVal) And tmpVal > 0 Then
        maxRowParam = CLng(tmpVal)
    Else
        maxRowParam = 0 ' auto detect per sheet
    End If

    ' workbook limits (depend on file format/version)
    workbookMaxCols = wsMaster.Columns.Count
    workbookMaxRows = wsMaster.Rows.Count

    ' if user-specified values exceed workbook limits, clamp and warn
    If maxColParam > 0 And maxColParam > workbookMaxCols Then
        MsgBox "Master!C2 (" & maxColParam & ") exceeds workbook column limit (" & workbookMaxCols & ")." & vbCrLf & _
               "Will use " & workbookMaxCols & " instead.", vbExclamation
        maxColParam = workbookMaxCols
    End If
    If maxRowParam > 0 And maxRowParam > workbookMaxRows Then
        MsgBox "Master!D2 (" & maxRowParam & ") exceeds workbook row limit (" & workbookMaxRows & ")." & vbCrLf & _
               "Will use " & workbookMaxRows & " instead.", vbExclamation
        maxRowParam = workbookMaxRows
    End If

    ' output file on Desktop
    outputFilePath = Environ("USERPROFILE") & "\ScanMarkersOutput.txt"
    fileNum = FreeFile
    Open outputFilePath For Output As #fileNum

    ' print header and workbook limits
    Print #fileNum, "==== Scan Start ===="
    Print #fileNum, "Workbook maximum rows: " & workbookMaxRows
    Print #fileNum, "Workbook maximum columns: " & workbookMaxCols
    Print #fileNum, ""

    ' loop sheets
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "Master" Then
            Print #fileNum, "---- Sheet: " & ws.Name & " ----"
            matchFound = False

            ' determine used range boundaries
            If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
                ' sheet empty: used range is at least row 1 col 1
                usedFirstRow = 1
                usedLastRow = 1
                usedFirstCol = 1
                usedLastCol = 1
            Else
                usedFirstRow = ws.UsedRange.Row
                usedLastRow = usedFirstRow + ws.UsedRange.Rows.Count - 1
                usedFirstCol = ws.UsedRange.Column
                usedLastCol = usedFirstCol + ws.UsedRange.Columns.Count - 1
            End If

            ' decide scanLastRow / scanLastCol
            If maxRowParam > 0 Then
                scanLastRow = maxRowParam
            Else
                scanLastRow = usedLastRow
            End If

            If maxColParam > 0 Then
                scanLastCol = maxColParam
            Else
                scanLastCol = usedLastCol
            End If

            ' clamp to sheet limits
            If scanLastRow < 1 Then scanLastRow = 1
            If scanLastRow > ws.Rows.Count Then
                scanLastRow = ws.Rows.Count
            End If
            If scanLastCol < 1 Then scanLastCol = 1
            If scanLastCol > ws.Columns.Count Then
                scanLastCol = ws.Columns.Count
            End If

            ' print scanning range
            Print #fileNum, "Scanning Rows 1 - " & scanLastRow & ", Cols 1 - " & scanLastCol

            ' scan cells
            For r = 1 To scanLastRow
                For c = 1 To scanLastCol
                    val = Trim(CStr(ws.Cells(r, c).Value))
                    If val <> "" Then
                        For Each kw In keywords
                            If InStr(1, val, CStr(kw), vbTextCompare) > 0 Then
                                outputText = "Row=" & r & " Col=" & c & " -> " & val
                                Print #fileNum, outputText
                                matchFound = True
                                Exit For
                            End If
                        Next kw
                    End If
                Next c
            Next r

            If Not matchFound Then
                Print #fileNum, "(No matches found)"
            End If

            Print #fileNum, "---- End of Sheet: " & ws.Name & " ----"
            Print #fileNum, ""
        End If
    Next ws

    Print #fileNum, "==== Scan End ===="
    Close #fileNum

    MsgBox "Scan completed. Output saved to: " & outputFilePath, vbInformation
    Exit Sub

    Dim lastRow As Long
    Dim maxRow As Long, maxCol As Long
    Dim marker As String
    Dim outputFile As String

    On Error GoTo ErrHandler

    '--- set worksheet ---
    Set ws = ThisWorkbook.Sheets("Master")

    '--- read max row/col from C2 / D2 ---
    maxRow = CLng(ws.Cells(2, 3).Value) ' C2
    maxCol = CLng(ws.Cells(2, 4).Value) ' D2

    '--- validate row/col limits ---
    If maxRow < 1 Or maxRow > ws.Rows.Count Then
        MsgBox "Invalid max row (C2). Valid range: 1 - " & ws.Rows.Count, vbCritical
        Exit Sub
    End If
    If maxCol < 1 Or maxCol > ws.Columns.Count Then
        MsgBox "Invalid max column (D2). Valid range: 1 - " & ws.Columns.Count, vbCritical
        Exit Sub
    End If

    '--- prepare output file ---
    outputFile = ThisWorkbook.Path & Application.PathSeparator & "ScanMarkersOutput.txt"

    ' check if file already open
    If IsFileLocked(outputFile) Then
        MsgBox "Cannot write to output file: " & outputFile & vbCrLf & _
               "It may already be open. Please close it and try again.", vbCritical
        Exit Sub
    End If

    fileNum = FreeFile
    Open outputFile For Output As #fileNum

    '--- scan markers ---
    For r = 1 To maxRow
        For c = 1 To maxCol
            marker = Trim(CStr(ws.Cells(r, c).Value))
            If Len(marker) > 0 Then
                Print #fileNum, "Row " & r & ", Col " & c & ": " & marker
            End If
        Next c
    Next r

    Close #fileNum

    MsgBox "Scan completed successfully." & vbCrLf & _
           "Output file: " & outputFile, vbInformation
    Exit Sub

'--- error handler ---
ErrHandler:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum

    Dim errMsg As String
    If Err.Number <> 0 Then
        errMsg = "Error " & Err.Number & vbCrLf & _
                 "Source: " & Err.Source & vbCrLf & _
                 "Description: " & Err.Description
    Else
        errMsg = "Unexpected error trapped, but Err.Number = 0." & vbCrLf & _
                 "Possible cause: Output file already open or locked."
    End If

    MsgBox errMsg, vbCritical, "ScanMarkers Error"
End Sub

'--- helper function to detect locked file ---
Private Function IsFileLocked(filePath As String) As Boolean
    On Error Resume Next
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary Access Read Write Lock Read Write As #ff
    Close #ff
    If Err.Number <> 0 Then
        IsFileLocked = True
        Err.Clear
    Else
        IsFileLocked = False
    End If
    On Error GoTo 0
End Function


ErrHandler:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Function


