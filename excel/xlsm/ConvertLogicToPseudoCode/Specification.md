规范文档
=====

---

# 📄 VBA 宏文档：ExtractValidationFlows

## 📌 功能目标

从工作表 **FunctionalSpecifications** 中扫描验证逻辑，自动提取 **Validation Flows**，并将结果输出到新工作表 **ValidationFlows**。

---

## 📑 输入格式规则

源表：`FunctionalSpecifications`

* **流程起点**

    * 每个流程以某行的 **`IF`** 开头（位于某一列，记作 *startCol*）。
    * 下一行的同一列必须是 **`Y`** → 才判定为流程开始。

* **流程结束条件**

    * 当在 *startCol* 列遇到 **`N`** 时结束，或扫描到表格结尾。

* **流程内容**

    * **顶层条件**：流程起点 `IF` 行的右侧内容。
    * **嵌套条件**：流程内部，*startCol+1..maxCol* 范围内的任意 `IF`。
    * **终止条件**：流程中出现 `[message].messageId` → 该行记作消息行。
    * 仅在找到 **messageId** 时，输出整个流程。

---

## 📤 输出格式规则

目标表：`ValidationFlows`

* 每个流程输出如下格式：

  ```
  Validation Flow <序号>
  IF <顶层条件>
  IF <嵌套条件1>
  IF <嵌套条件2>
  ...
  - [message].messageId = <值>
  ```
* 各流程之间空一行。

---

## 🛠️ 核心逻辑步骤

1. **确定范围**

    * 计算 `lastRow`（扫描多列，取最大行号）。
    * 设置 `maxCol = 50`（可调整）。

2. **逐行扫描**

    * 查找行内第一个 marker (`IF/Y/N`) → 定位 *startCol*。
    * 如果 `IF` + 下一行同列是 `Y` → 进入流程解析。

3. **解析流程**

    * 保存顶层 `IF`。
    * 向下逐行：

        * 遇到同列 `N` → 结束流程。
        * 检查 *startCol+1..maxCol* 是否有 `IF` → 记录为嵌套条件。
        * 查找 `[message].messageId` → 找到时记为消息行并结束解析。

4. **结果输出**

    * 按输出规则写入 `ValidationFlows` 表。

---

## 📦 宏接口说明

* 主过程：`Sub ExtractValidationFlows()`
* 辅助函数：

    * `GetRowText(ws, rowNum, colStart, colEnd)` → 拼接一行的内容
    * `FindFirstMarkerCol(ws, rowNum, maxCol)` → 定位 marker 列

---

## 📋 使用说明

1. 在 Excel 打开 VBA 编辑器 (Alt + F11)。
2. 新建模块，粘贴宏代码。
3. 确保存在源表 `FunctionalSpecifications`。
4. 运行 `ExtractValidationFlows`。
5. 结果会出现在新建/覆盖的 `ValidationFlows` 表。

---

## ✅ 输出示例

输入表：

```
Row2:   IF x <> BLANK
Row3:   IF Data does not exist == TRUE
Row4:   IF [y] <> [z]
Row5:   - [message].messageId = E10001
Row8:   IF y <> BLANK
Row9:   IF Data does not exist == TRUE
Row10:  - [message].messageId = E10002
```

输出表：

```
Validation Flow 1
IF x <> BLANK
IF Data does not exist == TRUE
IF [y] <> [z]
- [message].messageId = E10001

Validation Flow 2
IF y <> BLANK
IF Data does not exist == TRUE
- [message].messageId = E10002
```
