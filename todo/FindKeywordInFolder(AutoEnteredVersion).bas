Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindKeywordInFolder
' This script searches for a keyword in all Excel files within
' a selected folder and its subfolders, using a list of
' keywords from the "Keywords" sheet in the current workbook.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindKeywordInFolder()
    ' Variable Declarations
    Dim KeywordSheet As Worksheet
    Dim SearchKeyword As Variant ' Use Variant for loop
    Dim FolderPath As String
    Dim ReportWorkbook As Workbook
    Dim ReportSheet As Worksheet
    Dim nextRow As Long
    Dim FileSystemObj As Object
    Dim folder As Object

    Dim LastRow As Long
    Dim KeyWordCell As Range
    Dim KeywordsFoundCount As Long

    ' 1. Get Folder Path from User
    FolderPath = SelectFolder()

    ' Exit if user cancels the folder selection
    If FolderPath = "" Then Exit Sub

    ' 2. Identify the Keyword Sheet
    On Error GoTo KeywordSheetErrorHandler
    ' *** ASSUMPTION: Keywords are in a sheet named "Keywords" in the active workbook ***
    Set KeywordSheet = ThisWorkbook.Sheets("Keywords")
    On Error GoTo 0 ' Resume normal error handling

    ' 3. Create a new workbook for reporting results
    Set ReportWorkbook = Workbooks.Add
    Set ReportSheet = ReportWorkbook.Sheets(1)

    ' Clear previous results and set headers
    ReportSheet.Cells.Clear
    ReportSheet.Range("A1:F1").Value = Array("Keyword", "File Path", "Sheet Name", "Location Type", "Location Detail", "Found Text")
    ReportSheet.Range("A1:F1").Font.Bold = True
    nextRow = 2 ' Start reporting results from row 2

    ' 4. Create FileSystemObject to traverse folders (once)
    Set FileSystemObj = CreateObject("Scripting.FileSystemObject")
    Set folder = FileSystemObj.GetFolder(FolderPath)

    ' 5. Loop through the list of keywords in Column A
    With KeywordSheet
        ' Find the last row in Column A containing data
        LastRow = .Cells(.Rows.Count, "A").End(xlUp).Row

        If LastRow < 1 Or IsEmpty(.Range("A1").Value) Then
            MsgBox "No keywords found in Column A of the 'Keywords' sheet.", vbCritical
            GoTo FinalCleanup
        End If

        ' Loop through each cell in Column A
        For Each KeyWordCell In .Range("A1:A" & LastRow)
            SearchKeyword = Trim(KeyWordCell.Value)

            ' Skip if the keyword is empty
            If SearchKeyword <> "" Then
                ' Process all Excel files in the folder and subfolders for the current keyword
                ProcessFolder folder, CStr(SearchKeyword), ReportSheet, nextRow, FileSystemObj
            End If
        Next KeyWordCell
    End With

    KeywordsFoundCount = nextRow - 2

    ' Final Cleanup and Formatting
    ReportSheet.Columns("A:F").AutoFit

    ' Confirmation message
    MsgBox "Keyword search complete. Total matches found for all keywords: " & KeywordsFoundCount, vbInformation

    GoTo FinalCleanup

KeywordSheetErrorHandler:
    MsgBox "Could not find a sheet named 'Keywords' in this workbook. Please create it and list your search terms in Column A.", vbCritical
    Exit Sub

FinalCleanup:
    Set FileSystemObj = Nothing
    Set folder = Nothing
    Set ReportWorkbook = Nothing
    Set ReportSheet = Nothing
    Set KeywordSheet = Nothing
    Set KeyWordCell = Nothing
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcessFolder
' Recursively processes all files in a folder and its subfolders
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub ProcessFolder(ByVal folder As Object, ByVal SearchKeyword As String, ByVal ReportSheet As Worksheet, ByRef nextRow As Long, ByVal FileSystemObj As Object)
    Dim file As Object
    Dim subFolder As Object

    ' Process files in the current folder
    For Each file In folder.Files
        ' Check if the file is an Excel file
        If IsExcelFile(file.Name) Then
            SearchInWorkbook file.Path, SearchKeyword, ReportSheet, nextRow
        End If
    Next file

    ' Recursively process subfolders
    For Each subFolder In folder.SubFolders
        ProcessFolder subFolder, SearchKeyword, ReportSheet, nextRow, FileSystemObj
    Next subFolder
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsExcelFile
' Checks if a file is an Excel file based on its extension
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function IsExcelFile(FileName As String) As Boolean
    Dim ext As String
    ext = LCase(Right(FileName, 4))
    If ext = ".xls" Or ext = ".xlsx" Or ext = ".xlsm" Or ext = ".xlsb" Then
        IsExcelFile = True
    Else
        ext = LCase(Right(FileName, 5))
        If ext = ".xltx" Or ext = ".xltm" Then
            IsExcelFile = True
        Else
            IsExcelFile = False
        End If
    End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SearchInWorkbook
' Searches for the keyword in a specific workbook
' *** MODIFIED: Includes the keyword in the report output ***
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub SearchInWorkbook(FilePath As String, SearchKeyword As String, ReportSheet As Worksheet, ByRef nextRow As Long)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim cellMatch As Range
    Dim shp As Shape
    Dim shapeText As String
    Dim firstAddress As String
    Dim openedWorkbook As Boolean

    On Error GoTo ErrorHandler

    ' Try to get the workbook if it's already open
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(GetFileNameFromPath(FilePath))
    On Error GoTo ErrorHandler

    ' If the workbook isn't open, open it temporarily
    If wb Is Nothing Then
        Set wb = Workbooks.Open(FilePath, ReadOnly:=True, UpdateLinks:=0)
        openedWorkbook = True
    End If

    ' Search in each worksheet
    For Each ws In wb.Worksheets
        ' ----------------------------------------------------
        ' PART A: Search CELLS for the keyword (Fuzzy Match)
        ' ----------------------------------------------------

        With ws.usedRange
            ' Set up the initial search
            Set cellMatch = .Find(What:=SearchKeyword, _
                                 LookIn:=xlValues, _
                                 LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, _
                                 MatchCase:=False)

            If Not cellMatch Is Nothing Then
                firstAddress = cellMatch.Address
                Do
                    ' Record the match
                    ReportSheet.Cells(nextRow, "A").Value = SearchKeyword ' Added Keyword
                    ReportSheet.Cells(nextRow, "B").Value = FilePath
                    ReportSheet.Cells(nextRow, "C").Value = ws.Name
                    ReportSheet.Cells(nextRow, "D").Value = "Cell"
                    ReportSheet.Cells(nextRow, "E").Value = cellMatch.Address(External:=False)
                    ReportSheet.Cells(nextRow, "F").Value = cellMatch.Value ' Moved Found Text to column F
                    nextRow = nextRow + 1

                    ' Find the next match
                    Set cellMatch = .FindNext(cellMatch)
                Loop While Not cellMatch Is Nothing And cellMatch.Address <> firstAddress
            End If
        End With

        ' ----------------------------------------------------
        ' PART B: Search SHAPES for the keyword (Fuzzy Match)
        ' ----------------------------------------------------

        For Each shp In ws.Shapes
            shapeText = ""

            On Error Resume Next
            shapeText = shp.TextFrame2.TextRange.Text

            If Err.Number <> 0 Then
                Err.Clear
                GoTo NextShape
            End If

            On Error GoTo 0 ' Resume normal error handling

            If InStr(1, shapeText, SearchKeyword, vbTextCompare) > 0 Then
                ' Record the shape match
                ReportSheet.Cells(nextRow, "A").Value = SearchKeyword ' Added Keyword
                ReportSheet.Cells(nextRow, "B").Value = FilePath
                ReportSheet.Cells(nextRow, "C").Value = ws.Name
                ReportSheet.Cells(nextRow, "D").Value = "Shape"
                ReportSheet.Cells(nextRow, "E").Value = shp.Name
                ReportSheet.Cells(nextRow, "F").Value = Replace(shapeText, Chr(10), " ") ' Moved Found Text to column F
                nextRow = nextRow + 1
            End If

NextShape:
        Next shp
    Next ws

Cleanup:
    ' Close the workbook if we opened it
    If openedWorkbook Then
        wb.Close SaveChanges:=False
    End If

    Exit Sub

ErrorHandler:
    ' Handle errors
    ReportSheet.Cells(nextRow, "A").Value = SearchKeyword ' Record keyword with error
    ReportSheet.Cells(nextRow, "B").Value = FilePath & " [Error accessing file]"
    nextRow = nextRow + 1
    Resume Cleanup
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' GetFileNameFromPath
' Extracts file name from full file path
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function GetFileNameFromPath(FilePath As String) As String
    Dim parts
    parts = Split(FilePath, "\")
    GetFileNameFromPath = parts(UBound(parts))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SelectFolder
' Opens a folder picker dialog and returns the selected path
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function SelectFolder() As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)

    With fd
        .Title = "Select Folder to Search"
        .AllowMultiSelect = False
        If .Show = -1 Then
            SelectFolder = .SelectedItems(1)
        Else
            SelectFolder = ""
        End If
    End With
End Function

