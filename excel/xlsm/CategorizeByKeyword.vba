Sub CategorizeByKeyword()
     ' Declares all necessary variables
     Dim wsInput As Worksheet
     Dim lastRow As Long
     Dim i As Long
     Dim cellValue As String
     Dim keywords() As String
     Dim keyword As Variant
     Dim sheetExists As Boolean
     Dim targetSheet As Worksheet
     Dim targetRow As Long

     ' Set the source worksheet (assuming the input data is in a sheet named "Input")
     Set wsInput = ThisWorkbook.Sheets("Input")

     ' Find the last row with data in column D of the Input sheet
     lastRow = wsInput.Cells(wsInput.Rows.Count, 4).End(xlUp).Row

     ' Loop through each row of the Input sheet, starting from row 2 (assuming row 1 is a header)
     For i = 2 To lastRow
         cellValue = Trim(wsInput.Cells(i, 4).Value) ' Get the value from the cell in column D

         ' Check if the cell in column D is not empty
         If cellValue <> "" Then
             ' Split the cell value by line breaks to handle multiple keywords in one cell
             keywords = Split(cellValue, vbLf)

             ' Loop through each individual keyword
             For Each keyword In keywords
                 keyword = Trim(keyword) ' Remove any leading or trailing spaces from the keyword
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

                     ' Copy the content from columns E and F of the Input sheet to the target sheet
                     wsInput.Cells(i, 5).Copy Destination:=targetSheet.Cells(targetRow, 1) ' Copies column E content to column A
                     wsInput.Cells(i, 6).Copy Destination:=targetSheet.Cells(targetRow, 2) ' Copies column F content to column B
                 End If
             Next keyword
         End If
     Next i

     ' Inform the user that the process is complete
     MsgBox "Data categorization by keyword is complete!", vbInformation
 End Sub
