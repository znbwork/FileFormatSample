''''''''''''''''''''''''''''''''''''''''''''''''''
' ConsolidateData
''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ConsolidateData()
    ' Define the names of the settings and summary sheets.
    ' You can change these if your sheet names are different.
    Const SETTINGS_SHEET_NAME As String = "Settings"
    Const SUMMARY_SHEET_NAME As String = "Summary"

    Dim settingsWs As Worksheet
    Dim summaryWs As Worksheet
    Dim lastSettingsRow As Long
    Dim i As Long
    Dim sourceSheetName As String
    Dim colRangeString As String
    Dim startCol As Long
    Dim endCol As Long
    Dim colCounter As Long
    Dim sourceWs As Worksheet
    Dim lastSourceRow As Long
    Dim lastSummaryRow As Long

    ' --- Step 1: Set up the Summary Sheet ---

    ' Check if the Settings sheet exists.
    On Error GoTo SettingsSheetNotFound
    Set settingsWs = ThisWorkbook.Sheets(SETTINGS_SHEET_NAME)
    On Error GoTo 0

    ' Check if the Summary sheet already exists. If it does, clear it.
    ' If not, create a new one.
    On Error Resume Next
    Set summaryWs = ThisWorkbook.Sheets(SUMMARY_SHEET_NAME)
    On Error GoTo 0

    If summaryWs Is Nothing Then
        Set summaryWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        summaryWs.Name = SUMMARY_SHEET_NAME
    Else
        ' Clear all content from the existing Summary sheet.
        summaryWs.Cells.ClearContents
    End If

    ' --- Step 2: Loop through the Settings sheet to get consolidation instructions ---

    ' Find the last row with data in column A of the Settings sheet.
    lastSettingsRow = settingsWs.Cells(settingsWs.Rows.Count, "A").End(xlUp).Row

    ' Initialize a column counter for the Summary sheet.
    colCounter = 1

    ' Loop from the second row (A2) to the last row in the Settings sheet.
    For i = 2 To lastSettingsRow
        sourceSheetName = Trim(settingsWs.Cells(i, "A").Value)
        colRangeString = Trim(settingsWs.Cells(i, "B").Value)

        ' Skip if the sheet name or column range is empty.
        If sourceSheetName = "" Or colRangeString = "" Then
            GoTo NextIteration
        End If

        ' --- Step 3: Parse the column range and get the data ---

        ' Try to get the source worksheet object.
        On Error GoTo SourceSheetNotFound
        Set sourceWs = ThisWorkbook.Sheets(sourceSheetName)
        On Error GoTo 0

        ' Split the column range string by " to ".
        Dim parts() As String
        parts = Split(colRangeString, " to ")

        If UBound(parts) <> 1 Then
            MsgBox "Invalid column range format for sheet '" & sourceSheetName & "' in row " & i & ". Format should be like 'C to F' or '1 to 5'.", vbExclamation, "Error"
            GoTo NextIteration
        End If

        ' Convert column names (e.g., "C") or numbers to column numbers.
        On Error GoTo InvalidColumnRange
        If IsNumeric(parts(0)) Then
            startCol = CLng(parts(0))
        Else
            startCol = Columns(parts(0)).Column
        End If

        If IsNumeric(parts(1)) Then
            endCol = CLng(parts(1))
        Else
            endCol = Columns(parts(1)).Column
        End If
        On Error GoTo 0

        ' --- Step 4: Copy data from the source sheet to the summary sheet ---

        ' Find the last row with data in the source sheet.
        lastSourceRow = sourceWs.Cells.Find(What:="*", After:=sourceWs.Cells(1, 1), _
                                            SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

        ' Loop through each column in the specified range.
        Dim currentSourceCol As Long
        For currentSourceCol = startCol To endCol
            ' Check if there is data to copy.
            If lastSourceRow >= 1 Then
                ' Copy the entire column from the source sheet.
                sourceWs.Columns(currentSourceCol).Copy

                ' Paste the data into the next available column on the Summary sheet.
                summaryWs.Columns(colCounter).PasteSpecial xlPasteAll

                ' Get the last row on the summary sheet to ensure full paste
                lastSummaryRow = summaryWs.Cells(summaryWs.Rows.Count, colCounter).End(xlUp).Row

                ' Add a header for the newly pasted column, indicating the source sheet and column.
                summaryWs.Cells(1, colCounter).Value = sourceWs.Name & " - Col " & currentSourceCol

                ' Autofit the column for better readability.
                summaryWs.Columns(colCounter).AutoFit
            End If

            ' Increment the column counter for the next paste operation.
            colCounter = colCounter + 1
        Next currentSourceCol

NextIteration:
    Next i

    ' Clean up and finish.
    Application.CutCopyMode = False
    MsgBox "Data consolidation complete!", vbInformation, "Success"
    Exit Sub

' --- Error Handlers ---
SettingsSheetNotFound:
    MsgBox "Could not find the '" & SETTINGS_SHEET_NAME & "' sheet. Please make sure it exists and the name is correct.", vbExclamation, "Error"
    Exit Sub

SourceSheetNotFound:
    MsgBox "Could not find the source sheet named '" & sourceSheetName & "' from the Settings sheet. Please check the name.", vbExclamation, "Error"
    Resume NextIteration ' Continue to the next sheet in the list.

InvalidColumnRange:
    MsgBox "Invalid column specified for sheet '" & sourceSheetName & "' in row " & i & ". Please use a valid column number or letter (e.g., C).", vbExclamation, "Error"
    GoTo NextIteration ' Continue to the next sheet in the list.

End Sub