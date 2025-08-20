
### 使用步骤

1. 打开你的 `.xlsm` 文件（必须是启用宏的）。
2. 按 `Alt + F11` 打开 VBA 编辑器。
3. 在 `ThisWorkbook` 或者新建一个 `Module` 里粘贴下面代码。
4. 关闭 VBA 编辑器，回到 Excel。
5. 在「开发工具」里点击「宏」，运行 `ExtractErrorCheck`。

---

### VBA 脚本

```vba
Option Explicit

Sub ExtractErrorCheck()
    Dim wsItems As Worksheet
    Dim wsSource As Worksheet
    Dim wsResult As Worksheet
    Dim lastRowItems As Long, lastRowSource As Long
    Dim i As Long, j As Long
    Dim itemList As Collection
    Dim cellValue As String
    Dim item As Variant
    Dim found As Boolean
    Dim resultRow As Long
    
    ' 设置：项目一览放在 Sheet1 (A列)，设计书放在 Sheet2
    Set wsItems = ThisWorkbook.Sheets("Sheet1")
    Set wsSource = ThisWorkbook.Sheets("Sheet2")
    
    ' 新建或清空结果 Sheet
    On Error Resume Next
    Set wsResult = ThisWorkbook.Sheets("Result")
    If wsResult Is Nothing Then
        Set wsResult = ThisWorkbook.Sheets.Add
        wsResult.Name = "Result"
    Else
        wsResult.Cells.Clear
    End If
    On Error GoTo 0
    
    ' 读取项目一览到集合
    Set itemList = New Collection
    lastRowItems = wsItems.Cells(wsItems.Rows.Count, 1).End(xlUp).Row
    For i = 1 To lastRowItems
        If Trim(wsItems.Cells(i, 1).Value) <> "" Then
            itemList.Add Trim(wsItems.Cells(i, 1).Value)
        End If
    Next i
    
    ' 遍历设计书内容
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row
    resultRow = 1
    wsResult.Cells(resultRow, 1).Value = "Matched Rows"
    resultRow = resultRow + 1
    
    For i = 1 To lastRowSource
        cellValue = wsSource.Cells(i, 1).Value
        If cellValue <> "" Then
            found = False
            For Each item In itemList
                ' 查找 ".item名"
                If InStr(1, cellValue, "." & item, vbTextCompare) > 0 Then
                    found = True
                    Exit For
                End If
            Next item
            If found Then
                wsResult.Cells(resultRow, 1).Value = cellValue
                resultRow = resultRow + 1
            End If
        End If
    Next i
    
    MsgBox "处理完成！结果已输出到 [Result] 工作表。", vbInformation
End Sub
```

---

### 脚本说明

* **Sheet1**：放项目一览（`A列`，每行一个，如 `item1`、`item2`）。
* **Sheet2**：放英文设计书的内容（至少一列文本，通常在 `A列`）。
* **Result**：运行宏后会生成/清空这个表，把所有匹配的行输出。

---

## 📌 Python 实现步骤

### 1. 安装依赖

```bash
pip install pandas openpyxl
```

### 2. 示例Python脚本

假设：

* `items.xlsx` → `Sheet1` 里存放项目一览（`A列`：item1, item2, ...）
* `design.xlsx` → `Sheet2` 里存放英文设计书（`A列`：描述文本，比如 `tableA.item1 must be checked`）

代码如下：

```python
import pandas as pd
import re

# 文件路径
items_file = "items.xlsx"
design_file = "design.xlsx"
output_file = "result.xlsx"

# 读取项目一览
items_df = pd.read_excel(items_file, sheet_name="Sheet1", header=None)
items = items_df[0].dropna().astype(str).tolist()  # 转成list
pattern = re.compile(r"\b\w+\.(" + "|".join(map(re.escape, items)) + r")\b")

# 读取设计书
design_df = pd.read_excel(design_file, sheet_name="Sheet2", header=None)
design_df = design_df.fillna("")  # 避免 NaN

# 筛选匹配的行
matched_rows = []
for idx, row in design_df.iterrows():
    text = str(row[0])
    if pattern.search(text):
        matched_rows.append([text])

# 保存结果
result_df = pd.DataFrame(matched_rows, columns=["Matched Rows"])
result_df.to_excel(output_file, index=False)

print(f"✅ 处理完成，结果已输出到 {output_file}")
```

---

### 3. 脚本说明

* `items.xlsx`

  ```
  item1
  item2
  ```
* `design.xlsx`

  ```
  tableA.item1 must not be empty
  tableB.item2 should be unique
  tableC.item3 check length
  ```
* 运行后生成的 `result.xlsx` 内容：

  ```
  Matched Rows
  tableA.item1 must not be empty
  tableB.item2 should be unique
  ```

`tableC.item3` 被自动过滤掉，因为 `item3` 不在项目一览中。

---

### 4. 扩展功能

你还可以很容易加上：

* 输出 **匹配到的是哪个 item**（多列结果）。
* 扫描 Excel 的 **多列** 或 **多行**。
* 一次处理多个文档。

---

