Option Explicit

' Return concatenated text from colStart..colEnd for a given row (safe with empty cells)
Function GetRowText(ws As Worksheet, rowNum As Long, colStart As Long, colEnd As Long) As String
    Dim c As Long, s As String
    s = ""
    For c = colStart To colEnd
        On Error Resume Next
        Dim v As String
        v = Trim(CStr(ws.Cells(rowNum, c).Value & ""))
        On Error GoTo 0
        If v <> "" Then
            If s <> "" Then s = s & " "
            s = s & v
        End If
    Next c
    GetRowText = s
End Function

' Find the first marker column in the row that contains IF / Y / N (search left-to-right)
Function FindFirstMarkerCol(ws As Worksheet, rowNum As Long, maxCol As Long) As Long
    Dim c As Long, v As String
    For c = 1 To maxCol
        On Error Resume Next
        v = Trim(CStr(ws.Cells(rowNum, c).Value & ""))
        On Error GoTo 0
        If v <> "" Then
            Select Case UCase(v)
                Case "IF", "Y", "N"
                    FindFirstMarkerCol = c
                    Exit Function
            End Select
        End If
    Next c
    FindFirstMarkerCol = 0
End Function

Sub ExtractValidationFlows()
    On Error GoTo ErrHandler
    Dim wsSrc As Worksheet, wsDst As Worksheet
    Dim lastRow As Long, r As Long, maxCol As Long
    Dim dstRow As Long, flowIndex As Long
    Dim col As Long, tmpLast As Long

    maxCol = 50 ' adjust if your data extends farther right

    Set wsSrc = ThisWorkbook.Sheets("FunctionalSpecifications")
    If wsSrc Is Nothing Then
        MsgBox "Sheet 'FunctionalSpecifications' not found!", vbCritical
        Exit Sub
    End If

    ' compute lastRow robustly across columns 1..maxCol
    lastRow = 0
    For col = 1 To maxCol
        tmpLast = wsSrc.Cells(wsSrc.Rows.Count, col).End(xlUp).row
        If tmpLast > lastRow Then lastRow = tmpLast
    Next col
    If lastRow < 1 Then
        MsgBox "No data found on sheet.", vbExclamation
        Exit Sub
    End If

    ' prepare destination sheet
    On Error Resume Next
    Set wsDst = ThisWorkbook.Sheets("ValidationFlows")
    If wsDst Is Nothing Then
        Set wsDst = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsDst.Name = "ValidationFlows"
    Else
        On Error Resume Next
        wsDst.Cells.Clear
        On Error GoTo ErrHandler
    End If
    On Error GoTo ErrHandler

    dstRow = 1
    flowIndex = 0

    r = 1
    Do While r <= lastRow
        Dim marker As Long
        marker = FindFirstMarkerCol(wsSrc, r, maxCol)

        If marker > 0 Then
            Dim markerVal As String
            markerVal = UCase(Trim(CStr(wsSrc.Cells(r, marker).Value & "")))

            If markerVal = "IF" Then
                ' check next physical marker on next row - ensure same column and 'Y'
                Dim nextMarker As Long
                nextMarker = FindFirstMarkerCol(wsSrc, r + 1, maxCol)
                If nextMarker > 0 And nextMarker = marker Then
                    If UCase(Trim(CStr(wsSrc.Cells(r + 1, nextMarker).Value & ""))) = "Y" Then
                        ' Candidate flow start at row r, column 'marker'
                        Dim startCol As Long
                        Dim topIfExpr As String
                        Dim scanRow As Long
                        Dim messageFound As Boolean
                        Dim messageLine As String
                        Dim nestedIfs As Collection

                        Set nestedIfs = New Collection
                        startCol = marker
                        topIfExpr = Trim(GetRowText(wsSrc, r, marker + 1, maxCol))

                        ' scan forward until N in same column or end of sheet
                        scanRow = r + 1
                        messageFound = False
                        Do While scanRow <= lastRow
                            ' if at startCol there is an N, end this candidate flow
                            Dim valAtStart As String
                            valAtStart = UCase(Trim(CStr(wsSrc.Cells(scanRow, startCol).Value & "")))
                            If valAtStart = "N" Then
                                Exit Do
                            End If

                            ' scan this row for any IF markers to the right of startCol
                            Dim sc As Long
                            For sc = startCol + 1 To maxCol
                                Dim cellVal As String
                                cellVal = UCase(Trim(CStr(wsSrc.Cells(scanRow, sc).Value & "")))
                                If cellVal = "IF" Then
                                    Dim nestedExpr As String
                                    nestedExpr = Trim(GetRowText(wsSrc, scanRow, sc + 1, maxCol))
                                    If nestedExpr <> "" Then
                                        nestedIfs.Add nestedExpr
                                    End If
                                End If
                            Next sc

                            ' always check full row text for messageId
                            Dim fullText As String
                            fullText = Trim(GetRowText(wsSrc, scanRow, 1, maxCol))
                            If InStr(1, fullText, "[message].messageId", vbTextCompare) > 0 Then
                                messageFound = True
                                messageLine = fullText
                                Exit Do
                            End If

                            scanRow = scanRow + 1
                        Loop

                        ' Only output when a messageId was found inside this candidate range
                        If messageFound Then
                            flowIndex = flowIndex + 1
                            Debug.Print "Flow found: startRow=" & r & " startCol=" & startCol & " endRow=" & scanRow & " messageRow=" & scanRow

                            wsDst.Cells(dstRow, 1).Value = "Validation Flow " & flowIndex
                            dstRow = dstRow + 1

                            If topIfExpr <> "" Then
                                wsDst.Cells(dstRow, 1).Value = "IF " & topIfExpr
                                dstRow = dstRow + 1
                            End If

                            Dim i As Long
                            For i = 1 To nestedIfs.Count
                                wsDst.Cells(dstRow, 1).Value = "IF " & nestedIfs(i)
                                dstRow = dstRow + 1
                            Next i

                            messageLine = Replace(messageLine, """", "")
                            messageLine = Trim(Replace(messageLine, "- ", ""))
                            wsDst.Cells(dstRow, 1).Value = "- " & messageLine
                            dstRow = dstRow + 1

                            dstRow = dstRow + 1 ' blank line

                            ' skip processed rows by advancing r to scanRow (outer loop will increment further)
                            r = scanRow
                        End If

                        ' clean up
                        Set nestedIfs = Nothing
                    End If
                End If
            End If
        End If

        r = r + 1
    Loop

    Debug.Print "Extraction finished. Flows found: " & flowIndex
    MsgBox "Validation flows extracted to sheet 'ValidationFlows' (found " & flowIndex & " flows).", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
End Sub
