Option Explicit

Sub DebugScanMarkers()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, c As Long
    Dim val As String

    '
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FunctionalSpecifications")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'FunctionalSpecifications' not found!", vbCritical
        Exit Sub
    End If

    '
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).row

    Debug.Print "---- Start Scan ----"

    For r = 1 To lastRow
        For c = 1 To 50   '
            val = UCase(Trim(ws.Cells(r, c).Value))
            If val = "IF" Or val = "Y" Or val = "N" Then
                Debug.Print "Row=" & r & " Col=" & c & " -> " & val
            End If
        Next c
    Next r

    Debug.Print "---- End Scan ----"

    MsgBox "Scan finished! Please check Immediate Window (Ctrl+G).", vbInformation
End Sub

