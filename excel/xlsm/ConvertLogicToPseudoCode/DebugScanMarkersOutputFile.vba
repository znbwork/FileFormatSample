Option Explicit

Sub DebugScanMarkersOutputFile()
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long, c As Long
    Dim val As String
    Dim fileNum As Integer
    Dim outputFilePath As String
    Dim outputText As String

    ' Set output file path to desktop
    outputFilePath = Environ("USERPROFILE") & "\Desktop\ScanMarkersOutput.txt"
    
    ' Error handling when setting worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("FunctionalSpecifications")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Sheet 'FunctionalSpecifications' not found!", vbCritical
        Exit Sub
    End If

    ' Find last row
    lastRow = ws.Cells(ws.Rows.Count, 5).End(xlUp).row

    ' Open file for output
    fileNum = FreeFile
    Open outputFilePath For Output As #fileNum
    Print #fileNum, "---- Start Scan ----"

    For r = 1 To lastRow
        For c = 1 To 50   ' Check columns 1 to 50
            val = UCase(Trim(ws.Cells(r, c).Value))
            If val = "IF" Or val = "Y" Or val = "N" Then
                outputText = "Row=" & r & " Col=" & c & " -> " & val
                Print #fileNum, outputText
            End If
        Next c
    Next r

    Print #fileNum, "---- End Scan ----"
    Close #fileNum

    MsgBox "Scan finished! Output saved to: " & outputFilePath, vbInformation
End Sub