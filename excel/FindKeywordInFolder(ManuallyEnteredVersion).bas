Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' FindKeywordInFolder
' This script searches for a keyword in all Excel files within
' a selected folder and its subfolders.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Sub FindKeywordInFolder()
    ' Variable Declarations
    Dim SearchKeyword As String
    Dim FolderPath As String
    Dim ReportWorkbook As Workbook
    Dim ReportSheet As Worksheet
    Dim nextRow As Long
    Dim FileSystemObj As Object
    Dim folder As Object ' Explicitly declared as Object
    Dim file As Object   ' Explicitly declared as Object

    ' 1. Get Search Keyword from User
    SearchKeyword = Application.InputBox(Prompt:="Enter the keyword to search for (fuzzy match).", _
                                         Title:="Keyword Search", Type:=2)

    ' Exit if user cancels the input box
    If SearchKeyword = "False" Or SearchKeyword = "" Then Exit Sub

    ' 2. Get Folder Path from User
    FolderPath = SelectFolder()

    ' Exit if user cancels the folder selection
    If FolderPath = "" Then Exit Sub

    ' 3. Create a new workbook for reporting results
    Set ReportWorkbook = Workbooks.Add
    Set ReportSheet = ReportWorkbook.Sheets(1)

    ' Clear previous results and set headers
    ReportSheet.Cells.Clear
    ReportSheet.Range("A1:E1").Value = Array("File Path", "Sheet Name", "Location Type", "Location Detail", "Found Text")
    ReportSheet.Range("A1:E1").Font.Bold = True
    nextRow = 2 ' Start reporting results from row 2

    ' 4. Create FileSystemObject to traverse folders
    Set FileSystemObj = CreateObject("Scripting.FileSystemObject")
    Set folder = FileSystemObj.GetFolder(FolderPath)

    ' 5. Process all Excel files in the folder and subfolders
    ' All arguments are passed ByRef implicitly unless specified, except for Objects.
    ProcessFolder folder, SearchKeyword, ReportSheet, nextRow, FileSystemObj

    ' Final Cleanup and Formatting
    ReportSheet.Columns("A:E").AutoFit

    ' Confirmation message
    MsgBox "Keyword search complete. Total matches found: " & (nextRow - 2), vbInformation
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ProcessFolder
' Recursively processes all files in a folder and its subfolders
' *** FIX APPLIED: Explicitly declared all argument types ***
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

        ' Use Excel's built-in Find method for fast, fuzzy matching
        With ws.UsedRange
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
                    ReportSheet.Cells(nextRow, "A").Value = FilePath
                    ReportSheet.Cells(nextRow, "B").Value = ws.Name
                    ReportSheet.Cells(nextRow, "C").Value = "Cell"
                    ReportSheet.Cells(nextRow, "D").Value = cellMatch.Address(External:=False) ' Cell address
                    ReportSheet.Cells(nextRow, "E").Value = cellMatch.Value ' The cell content
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
            ' Reset text variable
            shapeText = ""

            ' Apply the robust error handling from the previous solution
            On Error Resume Next
            ' Note: This property only works for shapes with text frames (e.g., text boxes, certain AutoShapes)
            shapeText = shp.TextFrame2.TextRange.Text
            
            If Err.Number <> 0 Then
                Err.Clear
                ' If error, skip to the next shape
                GoTo NextShape
            End If
            
            On Error GoTo 0 ' Resume normal error handling
            
            ' Perform the fuzzy keyword search on the shape text
            If InStr(1, shapeText, SearchKeyword, vbTextCompare) > 0 Then
                ' Record the shape match
                ReportSheet.Cells(nextRow, "A").Value = FilePath
                ReportSheet.Cells(nextRow, "B").Value = ws.Name
                ReportSheet.Cells(nextRow, "C").Value = "Shape"
                ReportSheet.Cells(nextRow, "D").Value = shp.Name ' The shape's name
                
                ' Write the shape's content (replacing line feeds)
                ReportSheet.Cells(nextRow, "E").Value = Replace(shapeText, Chr(10), " ")
                nextRow = nextRow + 1
            End If
            
NextShape:
        Next shp ' End of Shape Loop
    Next ws ' End of Worksheet Loop
    
Cleanup:
    ' Close the workbook if we opened it
    If openedWorkbook Then
        wb.Close SaveChanges:=False
    End If
    
    Exit Sub
    
ErrorHandler:
    ' Handle errors (e.g., password protected files, corrupted files)
    ReportSheet.Cells(nextRow, "A").Value = FilePath & " [Error accessing file]"
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