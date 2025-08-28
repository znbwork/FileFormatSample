''''''''''''''''''''''''''''''''''''''''''''''''''
' CategorizeByKeyword
''''''''''''''''''''''''''''''''''''''''''''''''''
Sub CategorizeByKeyword()
     ' Declares all necessary variables
     Dim wsInput As Worksheet
     Dim lastRow As Long
     Dim i As Long, j As Long
     Dim cellValue As String
     Dim keywords() As String
     Dim keyword As Variant
     Dim sheetExists As Boolean
     Dim targetSheet As Worksheet
     Dim targetRow As Long

     ' Define the column for keywords
     Dim keywordColumn As Long

     ' Define the row range for keywords
     Dim startRowKeywords As Long
     Dim endRowKeywords As Long

     ' Define the column range for data to be extracted
     Dim startColData As Long
     Dim endColData As Long

     ' --- SET YOUR RANGES HERE ---
     ' This is where you specify the column containing keywords
     keywordColumn = 4 ' Column D

     ' This is where you specify the range of rows that contain your keywords.
     startRowKeywords = 7
     endRowKeywords = 10

     ' This is where you specify the columns you want to extract.
     startColData = 1  ' Column E
     endColData = 8    ' Column H
     ' --- END OF RANGE SETTINGS ---

     ' Set the source worksheet (assuming the input data is in a sheet named "Input")
     Set wsInput = ThisWorkbook.Sheets("Input")

     ' Find the last row with data in the entire sheet, to ensure all relevant rows are checked.
     ' Using UsedRange is a robust way to find the last row, regardless of empty cells.
     lastRow = wsInput.usedRange.Rows(wsInput.usedRange.Rows.Count).Row

     ' Loop through all rows in the Input sheet
     For i = 2 To lastRow
         ' Check if the current row falls within the specified keyword range
         If i >= startRowKeywords And i <= endRowKeywords Then
             ' Get the value from the cell in column D
             cellValue = Trim(wsInput.Cells(i, keywordColumn).Value)

             ' Check if the cell in column D is not empty
             If cellValue <> "" Then
                 ' Split the cell value by line breaks to handle multiple keywords
                 keywords = Split(cellValue, vbLf)

                 ' Loop through each individual keyword
                 For Each keyword In keywords
                     keyword = Trim(keyword) ' Remove any leading or trailing spaces
                     If keyword <> "" Then
                         ' Check if a worksheet with the keyword's name already exists
                         sheetExists = False
                         For Each targetSheet In ThisWorkbook.Worksheets
                             If targetSheet.Name = keyword Then
                                 sheetExists = True
                                 Set targetSheet = ThisWorkbook.Sheets(keyword)
                                 Exit For
                             End If
                         Next targetSheet

                         ' If the worksheet does not exist, create a new one
                         If Not sheetExists Then
                             Set targetSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                             targetSheet.Name = keyword
                         End If

                         ' Find the last row in the target sheet to append the new data
                         targetRow = targetSheet.Cells(targetSheet.Rows.Count, 1).End(xlUp).Row + 1

                         ' Copy the specified column range from the Input sheet to the target sheet
                         ' The destination starts at column A of the new sheet.
                         wsInput.Range(wsInput.Cells(i, startColData), wsInput.Cells(i, endColData)).Copy _
                             Destination:=targetSheet.Cells(targetRow, 1)
                     End If
                 Next keyword
             End If
         End If
     Next i

     ' Inform the user that the process is complete
     MsgBox "Data categorization by keyword is complete!", vbInformation
 End Sub