Sub ConvertLogicToPseudoCode_EN()

    Dim ws As Worksheet
    ' !!! Please modify based on your actual sheet name !!!
    Set ws = ThisWorkbook.Sheets("Functional Specifications")

    Dim lastRow As Long
    ' Assume your logic starts around row 240, find the last row of the data area
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    Dim outputString As String
    Dim currentRow As Long
    Dim indentLevel As Long ' Used to control indentation
    indentLevel = 0

    Dim cellEValue As String        ' Value in column E (Logic keywords: IF, NOP, TRUE/FALSE)
    Dim cellPValue As String        ' Value in column P (Used to identify the first IF, if needed)
    Dim cellGValue As String        ' Value in column G (Action/Description: Get data, message)
    Dim startRow As Long

    ' !!! Modify the starting row number based, e.g., startRow = 240 !!!
    startRow = 240

    ' Loop through the rows of interest
    For currentRow = startRow To lastRow

        ' Clean the key cell values for the current row
        cellEValue = Trim(ws.Cells(currentRow, "E").Value)
        cellPValue = Trim(ws.Cells(currentRow, "P").Value)
        cellGValue = Trim(ws.Cells(currentRow, "G").Value)

        ' Ignore empty rows
        If cellEValue = "" And cellPValue = "" And cellGValue = "" Then
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 1. Process Logic End and Decrease Indent (NOP)
        ' ----------------------------------------------------
        If InStr(1, cellEValue, "NOP", vbTextCompare) > 0 Then
            ' NOP usually indicates the end of a logic block
            indentLevel = indentLevel - 1
            If indentLevel < 0 Then indentLevel = 0

            ' If an NOP is followed by another NOP, ignore
            If InStr(1, ws.Cells(currentRow + 1, "E").Value, "NOP", vbTextCompare) > 0 Then
                GoTo NextRow
            End If

            GoTo NextRow ' NOP itself does not output pseudo-code
        End If

        ' ----------------------------------------------------
        ' 2. Build the current line's indentation
        ' ----------------------------------------------------
        Dim currentIndent As String
        currentIndent = String(indentLevel * 4, " ")

        ' ----------------------------------------------------
        ' 3. Process IF Conditions and Increase Indent
        ' ----------------------------------------------------

        If InStr(1, cellEValue, "IF", vbTextCompare) > 0 Then
            ' Extract the full IF condition (usually a combination of columns E, F, G, H)
            Dim fullCondition As String
            fullCondition = Trim(ws.Cells(currentRow, "E").Value & " " & _
                              ws.Cells(currentRow, "F").Value & " " & _
                              ws.Cells(currentRow, "G").Value & " " & _
                              ws.Cells(currentRow, "H").Value)

            ' Identify and simplify the condition
            Dim simplifiedCondition As String
            If InStr(1, fullCondition, "Data does not exist", vbTextCompare) > 0 Then
                simplifiedCondition = "IF Data_NOT_EXIST"
            ElseIf InStr(1, fullCondition, "CheckCcfMaintenance", vbTextCompare) > 0 Then
                simplifiedCondition = "IF Maintenance_Check_Condition"
            Else
                ' Otherwise, use the original E and G column values as the IF keyword
                simplifiedCondition = "IF " & Trim(ws.Cells(currentRow, "E").Value & " " & ws.Cells(currentRow, "G").Value)
            End If

            outputString = outputString & currentIndent & simplifiedCondition & ":" & vbCrLf
            indentLevel = indentLevel + 1 ' Increase indent
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 4. Process TRUE / FALSE Branches
        ' ----------------------------------------------------
        If InStr(1, cellEValue, "TRUE", vbTextCompare) > 0 Or InStr(1, cellEValue, "FALSE", vbTextCompare) > 0 Then
            ' TRUE/FALSE usually indicates the branch of an IF
            outputString = outputString & currentIndent & cellEValue & " BRANCH:" & vbCrLf
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 5. Process Specific Actions (Get Data, Message Setting, Variable Assignment)
        ' ----------------------------------------------------
        If cellGValue <> "" Then
            Dim processedAction As String

            ' A. Error Message Extraction (Based on your rule)
            If InStr(1, cellGValue, "[message].messageId", vbTextCompare) > 0 Then
                ' Use InStr/Mid to extract content between double quotes
                Dim startQuote As Long
                Dim endQuote As Long
                startQuote = InStr(cellGValue, Chr(34)) ' Find the first quote
                If startQuote > 0 Then
                    endQuote = InStr(startQuote + 1, cellGValue, Chr(34)) ' Find the second quote
                    If endQuote > 0 Then
                        Dim messageId As String
                        messageId = Mid(cellGValue, startQuote + 1, endQuote - startQuote - 1)
                        processedAction = "SET_ERROR_MSG: " & messageId
                    Else
                        processedAction = cellGValue ' Extraction failed, use original value
                    End If
                End If

            ' B. Get data operation
            ElseIf InStr(1, cellGValue, "Get data from TABLE", vbTextCompare) > 0 Then
                processedAction = "CALL: " & cellGValue

            ' C. Assignment or check operation
            ElseIf InStr(1, cellGValue, "=", vbTextCompare) > 0 Or InStr(1, cellGValue, "BLANK", vbTextCompare) > 0 Then
                processedAction = "ASSIGN/CHECK: " & cellGValue

            ' D. Other operations
            Else
                processedAction = cellGValue
            End If

            outputString = outputString & currentIndent & processedAction & vbCrLf
        End If

NextRow:
    Next currentRow

    ' ----------------------------------------------------
    ' 6. Output Result
    ' ----------------------------------------------------
    ' Output the result to the Immediate Window (Ctrl + G)
    Debug.Print "--- Extracted Pseudo Code ---"
    Debug.Print outputString

End Sub