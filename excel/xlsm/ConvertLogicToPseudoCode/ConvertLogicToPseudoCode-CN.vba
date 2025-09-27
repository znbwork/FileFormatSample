Sub ConvertLogicToPseudoCode()

    Dim ws As Worksheet
    ' !!! 请根据您的实际工作表名称修改 !!!
    Set ws = ThisWorkbook.Sheets("功能规格")

    Dim lastRow As Long
    ' 假设您的逻辑从第 240 行左右开始，我们找到数据区的最后一行
    lastRow = ws.Cells(ws.Rows.Count, "E").End(xlUp).Row

    Dim outputString As String
    Dim currentRow As Long
    Dim indentLevel As Long ' 用于控制缩进
    indentLevel = 0

    Dim cellValue As String
    Dim conditionValue As String
    Dim actionValue As String
    Dim startRow As Long

    ' !!! 根据逻辑开始的行号修改，例如从 240 行开始 !!!
    startRow = 240

    ' 遍历您感兴趣的行范围
    For currentRow = startRow To lastRow

        ' 清理当前行的关键单元格值
        cellValue = Trim(ws.Cells(currentRow, "E").Value) ' 逻辑关键词列 (IF, NOP, TRUE/FALSE)
        conditionValue = Trim(ws.Cells(currentRow, "P").Value) ' P 列的值 (用于识别第一个IF)
        actionValue = Trim(ws.Cells(currentRow, "G").Value) ' 操作/描述列 (Get data, message)

        ' 忽略空行
        If cellValue = "" And conditionValue = "" And actionValue = "" Then
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 1. 处理逻辑结束和减少缩进 (NOP, EndIF)
        ' ----------------------------------------------------
        If InStr(1, cellValue, "NOP", vbTextCompare) > 0 Then
            ' NOP 通常代表一个逻辑块的结束
            indentLevel = indentLevel - 1
            If indentLevel < 0 Then indentLevel = 0

            ' 如果 NOP 后面跟着另一个 NOP，通常是 End-to-End
            If InStr(1, ws.Cells(currentRow + 1, "E").Value, "NOP", vbTextCompare) > 0 Then
                GoTo NextRow ' 忽略连续的 NOP
            End If

            GoTo NextRow ' NOP 自身不输出伪代码行
        End If

        ' ----------------------------------------------------
        ' 2. 构建当前行的缩进
        ' ----------------------------------------------------
        Dim currentIndent As String
        currentIndent = String(indentLevel * 4, " ")

        ' ----------------------------------------------------
        ' 3. 处理 IF 条件和增加缩进
        ' ----------------------------------------------------

        ' A. 检查是否是第一个 IF (P列空行下的第一个IF)
        If InStr(1, cellValue, "IF", vbTextCompare) > 0 Then
            ' 提取完整的 IF 条件（通常在 E, F, G, H 列组合）
            Dim fullCondition As String
            fullCondition = Trim(ws.Cells(currentRow, "E").Value & " " & _
                              ws.Cells(currentRow, "F").Value & " " & _
                              ws.Cells(currentRow, "G").Value & " " & _
                              ws.Cells(currentRow, "H").Value)

            ' 识别并简化条件
            Dim simplifiedCondition As String
            If InStr(1, fullCondition, "Data does not exist", vbTextCompare) > 0 Then
                simplifiedCondition = "IF Data_NOT_EXIST"
            ElseIf InStr(1, fullCondition, "CheckCcfMaintenance", vbTextCompare) > 0 Then
                simplifiedCondition = "IF Maintenance_Check_Condition"
            Else
                ' 否则使用原始的 E 列值作为 IF 关键字
                simplifiedCondition = "IF " & Trim(ws.Cells(currentRow, "E").Value & " " & ws.Cells(currentRow, "G").Value)
            End If

            outputString = outputString & currentIndent & simplifiedCondition & ":" & vbCrLf
            indentLevel = indentLevel + 1 ' 增加缩进
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 4. 处理 ELSE / TRUE / FALSE 分支
        ' ----------------------------------------------------
        If InStr(1, cellValue, "TRUE", vbTextCompare) > 0 Or InStr(1, cellValue, "FALSE", vbTextCompare) > 0 Then
            ' 通常 TRUE/FALSE 表示 IF 的分支，且下一行才是真正的操作
            ' 此时缩进已经提前增加了，直接输出分支名称
            outputString = outputString & currentIndent & cellValue & " BRANCH:" & vbCrLf
            GoTo NextRow
        End If

        ' ----------------------------------------------------
        ' 5. 处理具体操作 (Get Data, Message Setting, Variable Assignment)
        ' ----------------------------------------------------
        If actionValue <> "" Then
            Dim processedAction As String

            ' A. 错误信息提取 (根据您的规则)
            If InStr(1, actionValue, "[message].messageId", vbTextCompare) > 0 Then
                ' 使用正则表达式或 InStr/Mid 来提取双引号内的内容
                Dim startQuote As Long
                Dim endQuote As Long
                startQuote = InStr(actionValue, Chr(34)) ' 找到第一个双引号
                If startQuote > 0 Then
                    endQuote = InStr(startQuote + 1, actionValue, Chr(34)) ' 找到第二个双引号
                    If endQuote > 0 Then
                        Dim messageId As String
                        messageId = Mid(actionValue, startQuote + 1, endQuote - startQuote - 1)
                        processedAction = "SET_ERROR_MSG: " & messageId
                    Else
                        processedAction = actionValue ' 提取失败，使用原值
                    End If
                End If

            ' B. Get data 操作
            ElseIf InStr(1, actionValue, "Get data from TABLE", vbTextCompare) > 0 Then
                processedAction = "CALL: " & actionValue

            ' C. 赋值或校验操作
            ElseIf InStr(1, actionValue, "=", vbTextCompare) > 0 Or InStr(1, actionValue, "BLANK", vbTextCompare) > 0 Then
                processedAction = "ASSIGN/CHECK: " & actionValue

            ' D. 其他操作
            Else
                processedAction = actionValue
            End If

            outputString = outputString & currentIndent & processedAction & vbCrLf
        End If

NextRow:
    Next currentRow

    ' ----------------------------------------------------
    ' 6. 输出结果
    ' ----------------------------------------------------
    ' 将结果输出到立即窗口 (Ctrl + G 打开)
    Debug.Print "--- Extracted Pseudo Code ---"
    Debug.Print outputString

    ' 可选：将结果写入新的工作表
    ' Worksheets.Add().Name = "PseudoCode_Output"
    ' Worksheets("PseudoCode_Output").Range("A1").Value = outputString

End Sub