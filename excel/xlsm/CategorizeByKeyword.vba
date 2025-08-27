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

    ' Define the row range for keywords in column D
    Dim startRowKeywords As Long
    Dim endRowKeywords As Long

    ' Set the keyword range here. For example, D7 to D10.
    startRowKeywords = 7
    endRowKeywords = 10

    ' Set the source worksheet (assuming the input data is in a sheet named "Input")
    Set wsInput = ThisWorkbook.Sheets("Input")

    ' Find the last row with data in the entire sheet, to ensure all relevant rows are checked
    lastRow = wsInput.UsedRange.Rows(wsInput.UsedRange.Rows.Count).Row

    ' Loop through all rows in the Input sheet
    For i = 2 To lastRow
        cellValue = Trim(wsInput.Cells(i, 4).Value) ' Get the value from the cell in column D

        ' Check if the cell in column D is within the specified keyword range
        If i >= startRowKeywords And i <= endRowKeywords Then
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
        End If
    Next i

    MsgBox "Data categorization by keyword is complete!", vbInformation
End Sub