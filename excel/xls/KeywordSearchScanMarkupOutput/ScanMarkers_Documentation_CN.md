# VBA 脚本说明文档：ScanMarkers

## 概述
本 VBA 宏用于扫描 Excel 工作表中指定范围内的单元格，并将结果输出到文本文件。  
宏会从 **Master 工作表的 C2（最大行数）** 和 **D2（最大列数）** 读取扫描范围参数。

---

## 功能特性
- 从 **Master!C2** 读取最大行数。
- 从 **Master!D2** 读取最大列数。
- 在指定范围内扫描所有单元格。
- 将结果输出到与工作簿相同目录下的 `ScanMarkersOutput.txt` 文件。
- 检查输出文件是否已被占用，避免错误。
- 提供 **明确的错误信息**，包含错误编号、来源和描述。

---

## 工作原理

### 输入参数
- **C2（最大行数）** → 定义扫描的行数上限（不能超过 Excel 最大行数）。
- **D2（最大列数）** → 定义扫描的列数上限（不能超过 Excel 最大列数）。

### 执行流程
1. 验证输入参数是否合法（是否在 Excel 的行/列范围内）。
2. 检查输出文件是否已被其他程序占用。
3. 遍历指定范围内的每个单元格。
4. 将非空单元格内容写入输出文件，并标注所在行列号。
5. 扫描完成后，弹出提示并显示输出文件路径。

### 输出文件
结果会保存为 `ScanMarkersOutput.txt`，文件示例：
```
Row 1, Col 2: ExampleValue
Row 5, Col 3: [message].messageId
```

---

## 错误处理
- 如果 **C2 或 D2** 的值无效，脚本会立即停止并提示错误。
- 如果输出文件已被占用（例如在记事本中打开），脚本会提示关闭文件后重试。
- 任何意外错误都会显示如下信息：
```
Error <错误编号>
Source: <错误来源>
Description: <错误描述>
```

---

## 辅助函数：IsFileLocked
脚本包含一个辅助函数，用于检测文件是否被占用：
```vba
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
```
该函数可以避免在输出文件被其他程序占用时导致 VBA 出错。

---

## 注意事项
- Excel 最大行数：**1,048,576**。
- Excel 最大列数：**16,384（XFD 列）**。
- 如果 C2 或 D2 超过这些范围，脚本会提示错误并停止执行。
- 请在运行前确认输出文件未被其他程序打开。

---

## 后续改进方向
- 增加筛选功能（例如仅输出包含 `[]` 的标记）。
- 支持追加写入，而不是每次覆盖文件。
- 增加对多工作表的扫描支持，而不仅仅是 `Master` 表。

---
