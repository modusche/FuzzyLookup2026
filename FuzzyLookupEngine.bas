Attribute VB_Name = "FuzzyLookupEngine"
Option Explicit

' ============================================================
' Fuzzy Lookup Engine - Core matching and output logic
' ============================================================

Public Type FuzzyConfig
    LeftRange As Range
    RightRange As Range
    LeftMatchCols() As Long
    RightMatchCols() As Long
    Threshold As Double
    MaxMatches As Long
    Method As String
    OutputSheet As String
End Type

Public Sub RunFuzzyLookup(cfg As FuzzyConfig)
    Dim wsOut As Worksheet
    Dim leftRow As Long, rightRow As Long
    Dim leftVal As String, rightVal As String
    Dim score As Double
    Dim outRow As Long
    Dim colIdx As Long
    Dim leftCols As Long, rightCols As Long
    Dim c As Long
    Dim matchCount As Long
    Dim totalLeft As Long, totalRight As Long
    Dim pctDone As Double

    ' Results storage for sorting
    Dim results() As Variant
    Dim resultCount As Long
    Dim maxResults As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo Cleanup

    leftCols = cfg.LeftRange.Columns.Count
    rightCols = cfg.RightRange.Columns.Count
    totalLeft = cfg.LeftRange.Rows.Count - 1  ' minus header
    totalRight = cfg.RightRange.Rows.Count - 1

    ' Create or clear output sheet
    On Error Resume Next
    Set wsOut = ThisWorkbook.Sheets(cfg.OutputSheet)
    If wsOut Is Nothing Then
        Set wsOut = ActiveWorkbook.Sheets(cfg.OutputSheet)
    End If
    On Error GoTo Cleanup

    If wsOut Is Nothing Then
        Set wsOut = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        wsOut.Name = cfg.OutputSheet
    Else
        wsOut.Cells.Clear
    End If

    ' Write headers
    outRow = 1
    c = 1

    ' Left table headers
    For colIdx = 1 To leftCols
        wsOut.Cells(outRow, c).Value = "Left." & cfg.LeftRange.Cells(1, colIdx).Value
        c = c + 1
    Next colIdx

    ' Right table headers
    For colIdx = 1 To rightCols
        wsOut.Cells(outRow, c).Value = "Right." & cfg.RightRange.Cells(1, colIdx).Value
        c = c + 1
    Next colIdx

    ' Similarity score header
    wsOut.Cells(outRow, c).Value = "Similarity"
    wsOut.Rows(1).Font.Bold = True

    outRow = 2

    ' Pre-cache right table values for performance
    Dim rightCache() As String
    ReDim rightCache(1 To totalRight, LBound(cfg.RightMatchCols) To UBound(cfg.RightMatchCols))
    For rightRow = 1 To totalRight
        For colIdx = LBound(cfg.RightMatchCols) To UBound(cfg.RightMatchCols)
            rightCache(rightRow, colIdx) = CStr(cfg.RightRange.Cells(rightRow + 1, cfg.RightMatchCols(colIdx)).Value)
        Next colIdx
    Next rightRow

    ' Main matching loop
    Dim statusBar As String
    For leftRow = 1 To totalLeft
        ' Status update every 10 rows
        If leftRow Mod 10 = 0 Or leftRow = totalLeft Then
            pctDone = CDbl(leftRow) / CDbl(totalLeft) * 100
            Application.StatusBar = "Fuzzy Lookup: Processing row " & leftRow & " of " & totalLeft & _
                                    " (" & Format(pctDone, "0") & "%)"
            DoEvents
        End If

        ' Collect matches for this left row
        matchCount = 0
        maxResults = cfg.MaxMatches * 2  ' over-allocate
        If maxResults < 10 Then maxResults = 10
        ReDim results(1 To maxResults, 1 To 2)  ' (rightRow, score)

        For rightRow = 1 To totalRight
            ' Calculate combined score across all match column pairs
            score = 0
            For colIdx = LBound(cfg.LeftMatchCols) To UBound(cfg.LeftMatchCols)
                leftVal = CStr(cfg.LeftRange.Cells(leftRow + 1, cfg.LeftMatchCols(colIdx)).Value)
                rightVal = rightCache(rightRow, colIdx)
                score = score + CombinedSimilarity(leftVal, rightVal, cfg.Method)
            Next colIdx
            ' Average across column pairs
            score = score / (UBound(cfg.LeftMatchCols) - LBound(cfg.LeftMatchCols) + 1)

            If score >= cfg.Threshold Then
                matchCount = matchCount + 1
                If matchCount > UBound(results, 1) Then
                    ReDim Preserve results(1 To matchCount + 10, 1 To 2)
                End If
                results(matchCount, 1) = rightRow
                results(matchCount, 2) = score
            End If
        Next rightRow

        ' Sort results by score descending (simple bubble sort, small n)
        If matchCount > 1 Then
            Dim ii As Long, jj As Long
            Dim tmpRow As Variant, tmpScore As Variant
            For ii = 1 To matchCount - 1
                For jj = 1 To matchCount - ii
                    If results(jj, 2) < results(jj + 1, 2) Then
                        tmpRow = results(jj, 1): tmpScore = results(jj, 2)
                        results(jj, 1) = results(jj + 1, 1): results(jj, 2) = results(jj + 1, 2)
                        results(jj + 1, 1) = tmpRow: results(jj + 1, 2) = tmpScore
                    End If
                Next jj
            Next ii
        End If

        ' Output top N matches
        Dim outputCount As Long
        outputCount = matchCount
        If outputCount > cfg.MaxMatches Then outputCount = cfg.MaxMatches

        For ii = 1 To outputCount
            c = 1
            rightRow = CLng(results(ii, 1))

            ' Left table values
            For colIdx = 1 To leftCols
                wsOut.Cells(outRow, c).Value = cfg.LeftRange.Cells(leftRow + 1, colIdx).Value
                c = c + 1
            Next colIdx

            ' Right table values
            For colIdx = 1 To rightCols
                wsOut.Cells(outRow, c).Value = cfg.RightRange.Cells(rightRow + 1, colIdx).Value
                c = c + 1
            Next colIdx

            ' Similarity score
            wsOut.Cells(outRow, c).Value = Round(results(ii, 2), 4)
            wsOut.Cells(outRow, c).NumberFormat = "0.00%"

            outRow = outRow + 1
        Next ii
    Next leftRow

    ' Autofit columns
    wsOut.UsedRange.Columns.AutoFit

    ' Format as table if possible
    On Error Resume Next
    If outRow > 2 Then
        wsOut.ListObjects.Add(xlSrcRange, wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(outRow - 1, leftCols + rightCols + 1)), , xlYes).Name = "FuzzyResults"
    End If
    On Error GoTo Cleanup

    MsgBox "Fuzzy Lookup Complete!" & vbCrLf & vbCrLf & _
           "Processed: " & totalLeft & " rows" & vbCrLf & _
           "Matches found: " & (outRow - 2) & vbCrLf & _
           "Results in sheet: " & cfg.OutputSheet, vbInformation, "Fuzzy Lookup"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

    If Err.Number <> 0 Then
        MsgBox "Error during Fuzzy Lookup: " & Err.Description, vbCritical, "Fuzzy Lookup Error"
    End If
End Sub
