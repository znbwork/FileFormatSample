# VBA Script Documentation: ScanMarkers

## Overview
This VBA macro scans an Excel worksheet based on user-defined parameters and exports the results to a text file.  
It is designed to read **maximum row (C2)** and **maximum column (D2)** values from the `Master` sheet and scan all cells within this range.

---

## Features
- Reads configuration from **Master!C2 (max rows)** and **Master!D2 (max columns)**.
- Scans all cells within the given range.
- Exports results into `ScanMarkersOutput.txt` located in the same folder as the workbook.
- Detects whether the output file is already open/locked, preventing errors.
- Provides **clear error messages** including error number, source, and description.

---

## How It Works

### Input Parameters
- **C2 (Row limit)** → Defines how many rows to scan (up to the maximum supported by Excel).
- **D2 (Column limit)** → Defines how many columns to scan (up to Excel’s maximum).

### Process
1. Validate inputs (must be within Excel’s row/column limits).
2. Check if the output file is locked (already open).
3. Loop through each cell within the specified range.
4. Write non-empty cell contents to the output file with row/column coordinates.
5. Show a completion message with the output file path.

### Output
- Results are written to `ScanMarkersOutput.txt` in the workbook directory.
- Example output:
  ```
  Row 1, Col 2: ExampleValue
  Row 5, Col 3: [message].messageId
  ```

---

## Error Handling
- If **C2 or D2** contain invalid values, the macro will stop with a clear error message.
- If the output file is already open, the macro will ask the user to close it before retrying.
- Any unexpected error will show:
  ```
  Error <Number>
  Source: <Error Source>
  Description: <Error Description>
  ```

---

## Helper Function: IsFileLocked
The script includes a helper function:
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
This checks whether the output file is currently locked by another process (e.g., already open in Notepad).

---

## Notes
- Excel row limit (modern versions): **1,048,576**.
- Excel column limit: **16,384** (column XFD).
- The macro will validate and stop if inputs exceed these limits.
- Best practice: Keep the output file closed before running the script.

---

## Future Improvements
- Add filtering (e.g., only write cells containing markers like `[ ]`).
- Support appending instead of overwriting the output file.
- Optionally scan all sheets, not only the `Master` sheet.

---
