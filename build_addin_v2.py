"""
Build Fuzzy Lookup Add-In v2.1 for Excel 2024 (64-bit)
Replicates the original Microsoft Fuzzy Lookup task pane experience.

Fix: Controls are now added at DESIGN TIME via the Designer object,
so VBA event handlers (btnGo_Click, etc.) fire correctly.
"""
import win32com.client
import os
import sys
import time

ADDIN_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PATH = os.path.join(ADDIN_DIR, "FuzzyLookup.xlam")

# =====================================================================
# VBA MODULE: FuzzyAlgorithms
# =====================================================================
MOD_ALGORITHMS = """
Option Explicit

Public Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long, len2 As Long
    Dim matrix() As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim above As Long, leftV As Long, diag As Long

    s1 = LCase$(s1): s2 = LCase$(s2)
    len1 = Len(s1): len2 = Len(s2)

    If len1 = 0 Then LevenshteinDistance = len2: Exit Function
    If len2 = 0 Then LevenshteinDistance = len1: Exit Function

    ReDim matrix(0 To len1, 0 To len2)
    For i = 0 To len1: matrix(i, 0) = i: Next i
    For j = 0 To len2: matrix(0, j) = j: Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then cost = 0 Else cost = 1
            above = matrix(i - 1, j) + 1
            leftV = matrix(i, j - 1) + 1
            diag = matrix(i - 1, j - 1) + cost
            matrix(i, j) = above
            If leftV < matrix(i, j) Then matrix(i, j) = leftV
            If diag < matrix(i, j) Then matrix(i, j) = diag
        Next j
    Next i
    LevenshteinDistance = matrix(len1, len2)
End Function

Public Function LevenshteinSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim maxLen As Long
    maxLen = Len(s1): If Len(s2) > maxLen Then maxLen = Len(s2)
    If maxLen = 0 Then LevenshteinSimilarity = 1#: Exit Function
    LevenshteinSimilarity = 1# - (CDbl(LevenshteinDistance(s1, s2)) / CDbl(maxLen))
End Function

Public Function JaroSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim len1 As Long, len2 As Long, matchDist As Long
    Dim matches As Long, transpositions As Long
    Dim s1Matches() As Boolean, s2Matches() As Boolean
    Dim i As Long, j As Long, k As Long

    s1 = LCase$(s1): s2 = LCase$(s2)
    len1 = Len(s1): len2 = Len(s2)

    If len1 = 0 And len2 = 0 Then JaroSimilarity = 1#: Exit Function
    If len1 = 0 Or len2 = 0 Then JaroSimilarity = 0#: Exit Function

    If len1 > len2 Then matchDist = CLng(len1 / 2) - 1 Else matchDist = CLng(len2 / 2) - 1
    If matchDist < 0 Then matchDist = 0

    ReDim s1Matches(1 To len1): ReDim s2Matches(1 To len2)
    matches = 0

    For i = 1 To len1
        Dim startJ As Long, endJ As Long
        startJ = i - matchDist: If startJ < 1 Then startJ = 1
        endJ = i + matchDist: If endJ > len2 Then endJ = len2
        For j = startJ To endJ
            If Not s2Matches(j) And Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                s1Matches(i) = True: s2Matches(j) = True: matches = matches + 1: Exit For
            End If
        Next j
    Next i

    If matches = 0 Then JaroSimilarity = 0#: Exit Function

    k = 1: transpositions = 0
    For i = 1 To len1
        If s1Matches(i) Then
            Do While Not s2Matches(k): k = k + 1: Loop
            If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then transpositions = transpositions + 1
            k = k + 1
        End If
    Next i

    JaroSimilarity = (CDbl(matches) / CDbl(len1) + CDbl(matches) / CDbl(len2) + _
                      (CDbl(matches) - CDbl(transpositions) / 2#) / CDbl(matches)) / 3#
End Function

Public Function JaroWinklerSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim jaro As Double, prefixLen As Long, maxP As Long, i As Long
    jaro = JaroSimilarity(s1, s2)
    maxP = 4
    If Len(s1) < maxP Then maxP = Len(s1)
    If Len(s2) < maxP Then maxP = Len(s2)
    prefixLen = 0
    For i = 1 To maxP
        If LCase$(Mid$(s1, i, 1)) = LCase$(Mid$(s2, i, 1)) Then prefixLen = prefixLen + 1 Else Exit For
    Next i
    JaroWinklerSimilarity = jaro + (CDbl(prefixLen) * 0.1 * (1# - jaro))
End Function

Public Function JaccardSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim tokens1() As String, tokens2() As String
    Dim dict As Object, i As Long, token As String
    Dim unionCount As Long, intersectCount As Long

    s1 = LCase$(Trim$(s1)): s2 = LCase$(Trim$(s2))
    If Len(s1) = 0 And Len(s2) = 0 Then JaccardSimilarity = 1#: Exit Function
    If Len(s1) = 0 Or Len(s2) = 0 Then JaccardSimilarity = 0#: Exit Function

    tokens1 = Split(s1, " "): tokens2 = Split(s2, " ")
    Set dict = CreateObject("Scripting.Dictionary")

    For i = LBound(tokens1) To UBound(tokens1)
        token = Trim$(tokens1(i))
        If Len(token) > 0 And Not dict.Exists(token) Then dict.Add token, 1
    Next i
    For i = LBound(tokens2) To UBound(tokens2)
        token = Trim$(tokens2(i))
        If Len(token) > 0 Then
            If dict.Exists(token) Then dict(token) = 3 Else dict.Add token, 2
        End If
    Next i

    unionCount = dict.Count: intersectCount = 0
    Dim v As Variant
    For Each v In dict.Items: If v = 3 Then intersectCount = intersectCount + 1
    Next v
    If unionCount = 0 Then JaccardSimilarity = 0# Else JaccardSimilarity = CDbl(intersectCount) / CDbl(unionCount)
End Function

Public Function CombinedSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim lev As Double, jw As Double, jac As Double
    lev = LevenshteinSimilarity(s1, s2)
    jw = JaroWinklerSimilarity(s1, s2)
    jac = JaccardSimilarity(s1, s2)
    Dim best As Double
    best = lev: If jw > best Then best = jw: If jac > best Then best = jac
    CombinedSimilarity = best * 0.6 + (lev + jw + jac) / 3# * 0.4
End Function

' === Worksheet functions ===
Public Function FUZZYMATCH(ByVal value1 As String, ByVal value2 As String) As Double
    Application.Volatile False
    FUZZYMATCH = CombinedSimilarity(value1, value2)
End Function

Public Function FUZZYVLOOKUP(ByVal lookupValue As String, ByVal tableRange As Range, _
                              ByVal matchCol As Long, ByVal returnCol As Long, _
                              Optional ByVal threshold As Double = 0.6) As Variant
    Application.Volatile False
    Dim bestScore As Double, bestRow As Long, score As Double, i As Long
    bestScore = 0: bestRow = 0
    For i = 1 To tableRange.Rows.Count
        Dim cv As String: cv = CStr(tableRange.Cells(i, matchCol).Value)
        If Len(cv) > 0 Then
            score = CombinedSimilarity(lookupValue, cv)
            If score > bestScore Then bestScore = score: bestRow = i
        End If
    Next i
    If bestScore >= threshold And bestRow > 0 Then
        FUZZYVLOOKUP = tableRange.Cells(bestRow, returnCol).Value
    Else
        FUZZYVLOOKUP = CVErr(xlErrNA)
    End If
End Function
"""

# =====================================================================
# VBA MODULE: FuzzyEngine
# =====================================================================
MOD_ENGINE = """
Option Explicit

Public Type TransformRule
    FromText As String
    ToText As String
End Type

Public Function ApplyTransformations(ByVal text As String, rules() As TransformRule, ruleCount As Long) As String
    Dim i As Long, result As String
    result = text
    For i = 0 To ruleCount - 1
        If Len(rules(i).FromText) > 0 Then
            result = Replace(result, rules(i).FromText, rules(i).ToText, , , vbTextCompare)
        End If
    Next i
    ApplyTransformations = result
End Function

Public Function CleanText(ByVal text As String, trimSpaces As Boolean, toLower As Boolean, removePunct As Boolean) As String
    Dim result As String
    result = text
    If toLower Then result = LCase$(result)
    If removePunct Then
        Dim i As Long, ch As String, cleaned As String
        cleaned = ""
        For i = 1 To Len(result)
            ch = Mid$(result, i, 1)
            Select Case Asc(ch)
                Case 32, 48 To 57, 65 To 90, 97 To 122: cleaned = cleaned & ch
            End Select
        Next i
        result = cleaned
    End If
    If trimSpaces Then
        result = Trim$(result)
        Do While InStr(result, "  ") > 0: result = Replace(result, "  ", " "): Loop
    End If
    CleanText = result
End Function

Public Sub ExecuteFuzzyLookup(leftWs As Worksheet, leftAddr As String, leftMatchCol As Long, _
                               rightWs As Worksheet, rightAddr As String, rightMatchCol As Long, _
                               threshold As Double, maxMatches As Long, _
                               doTrim As Boolean, doLower As Boolean, doRemovePunct As Boolean, _
                               outputSheetName As String, _
                               rules() As TransformRule, ruleCount As Long)

    Dim leftRange As Range, rightRange As Range, wsOut As Worksheet
    Dim leftRow As Long, rightRow As Long
    Dim leftVal As String, rightVal As String
    Dim score As Double, outRow As Long, c As Long
    Dim leftCols As Long, rightCols As Long
    Dim totalLeft As Long, totalRight As Long
    Dim colIdx As Long
    Dim startTime As Single

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup

    startTime = Timer

    Set leftRange = leftWs.Range(leftAddr)
    Set rightRange = rightWs.Range(rightAddr)
    leftCols = leftRange.Columns.Count
    rightCols = rightRange.Columns.Count
    totalLeft = leftRange.Rows.Count - 1
    totalRight = rightRange.Rows.Count - 1

    If totalLeft < 1 Or totalRight < 1 Then
        MsgBox "Tables must have at least a header row and one data row.", vbExclamation, "Fuzzy Lookup"
        GoTo Cleanup
    End If

    ' === BULK READ all data into arrays (massive speed boost) ===
    Dim leftData As Variant, rightData As Variant
    leftData = leftRange.Value    ' 2D array (1-based)
    rightData = rightRange.Value

    ' Pre-clean match column values
    Dim leftCache() As String, rightCache() As String
    ReDim leftCache(1 To totalLeft)
    ReDim rightCache(1 To totalRight)

    Application.StatusBar = "Fuzzy Lookup: Preparing data..."
    DoEvents

    For rightRow = 1 To totalRight
        rightVal = CStr(rightData(rightRow + 1, rightMatchCol))
        rightVal = CleanText(rightVal, doTrim, doLower, doRemovePunct)
        If ruleCount > 0 Then rightVal = ApplyTransformations(rightVal, rules, ruleCount)
        rightCache(rightRow) = rightVal
    Next rightRow

    For leftRow = 1 To totalLeft
        leftVal = CStr(leftData(leftRow + 1, leftMatchCol))
        leftVal = CleanText(leftVal, doTrim, doLower, doRemovePunct)
        If ruleCount > 0 Then leftVal = ApplyTransformations(leftVal, rules, ruleCount)
        leftCache(leftRow) = leftVal
    Next leftRow

    ' === Collect all results into array, then bulk write ===
    Dim totalOutCols As Long
    totalOutCols = leftCols + rightCols + 1

    ' Estimate max output rows
    Dim maxOutRows As Long
    maxOutRows = totalLeft * maxMatches
    If maxOutRows > 1000000 Then maxOutRows = 1000000

    Dim outData() As Variant
    ReDim outData(1 To maxOutRows, 1 To totalOutCols)
    outRow = 0

    ' Main matching loop
    Dim matchScores() As Double, matchRows() As Long
    ReDim matchScores(1 To totalRight)
    ReDim matchRows(1 To totalRight)

    For leftRow = 1 To totalLeft
        If leftRow Mod 20 = 0 Or leftRow = totalLeft Then
            Application.StatusBar = "Fuzzy Lookup: " & leftRow & " / " & totalLeft & _
                                    " (" & Format(CDbl(leftRow) / CDbl(totalLeft) * 100, "0") & "%)"
            DoEvents
        End If

        leftVal = leftCache(leftRow)
        If Len(leftVal) = 0 Then GoTo NextLeftRow

        Dim matchCount As Long
        matchCount = 0

        For rightRow = 1 To totalRight
            If Len(rightCache(rightRow)) = 0 Then GoTo NextRightRow
            score = CombinedSimilarity(leftVal, rightCache(rightRow))
            If score >= threshold Then
                matchCount = matchCount + 1
                matchScores(matchCount) = score
                matchRows(matchCount) = rightRow
            End If
NextRightRow:
        Next rightRow

        ' Sort top matches descending
        If matchCount > 1 Then
            Dim ii As Long, jj As Long, tmpS As Double, tmpR As Long
            For ii = 2 To matchCount
                tmpS = matchScores(ii): tmpR = matchRows(ii): jj = ii - 1
                Do While jj >= 1
                    If matchScores(jj) < tmpS Then
                        matchScores(jj + 1) = matchScores(jj): matchRows(jj + 1) = matchRows(jj): jj = jj - 1
                    Else: Exit Do
                    End If
                Loop
                matchScores(jj + 1) = tmpS: matchRows(jj + 1) = tmpR
            Next ii
        End If

        Dim outputN As Long
        outputN = matchCount: If outputN > maxMatches Then outputN = maxMatches

        For ii = 1 To outputN
            outRow = outRow + 1
            If outRow > maxOutRows Then
                maxOutRows = maxOutRows + 10000
                ReDim Preserve outData(1 To maxOutRows, 1 To totalOutCols)
            End If
            c = 1
            For colIdx = 1 To leftCols
                outData(outRow, c) = leftData(leftRow + 1, colIdx)
                c = c + 1
            Next colIdx
            For colIdx = 1 To rightCols
                outData(outRow, c) = rightData(matchRows(ii) + 1, colIdx)
                c = c + 1
            Next colIdx
            outData(outRow, c) = Round(matchScores(ii), 4)
        Next ii

NextLeftRow:
    Next leftRow

    ' === Create output sheet ===
    On Error Resume Next
    Set wsOut = ActiveWorkbook.Sheets(outputSheetName)
    On Error GoTo Cleanup
    If wsOut Is Nothing Then
        Set wsOut = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        wsOut.Name = outputSheetName
    Else
        wsOut.Cells.Clear
    End If

    Application.StatusBar = "Fuzzy Lookup: Writing results..."
    DoEvents

    ' Write headers
    c = 1
    For colIdx = 1 To leftCols
        wsOut.Cells(1, c).Value = leftData(1, colIdx): c = c + 1
    Next colIdx
    For colIdx = 1 To rightCols
        wsOut.Cells(1, c).Value = rightData(1, colIdx): c = c + 1
    Next colIdx
    wsOut.Cells(1, c).Value = "Similarity"

    ' Style headers
    With wsOut.Range(wsOut.Cells(1, 1), wsOut.Cells(1, leftCols))
        .Interior.Color = RGB(68, 114, 196): .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With
    With wsOut.Range(wsOut.Cells(1, leftCols + 1), wsOut.Cells(1, leftCols + rightCols))
        .Interior.Color = RGB(47, 117, 181): .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With
    With wsOut.Cells(1, totalOutCols)
        .Interior.Color = RGB(0, 176, 80): .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With

    ' === BULK WRITE results (massive speed boost) ===
    If outRow > 0 Then
        Dim writeData() As Variant
        ReDim writeData(1 To outRow, 1 To totalOutCols)
        For ii = 1 To outRow
            For c = 1 To totalOutCols
                writeData(ii, c) = outData(ii, c)
            Next c
        Next ii
        wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(outRow + 1, totalOutCols)).Value = writeData

        ' Format similarity column
        wsOut.Range(wsOut.Cells(2, totalOutCols), wsOut.Cells(outRow + 1, totalOutCols)).NumberFormat = "0.00%"
        wsOut.Range(wsOut.Cells(2, totalOutCols), wsOut.Cells(outRow + 1, totalOutCols)).Font.Bold = True
    End If

    ' Borders
    On Error Resume Next
    If outRow > 0 Then
        wsOut.Range(wsOut.Cells(1, leftCols + 1), wsOut.Cells(outRow + 1, leftCols + 1)).Borders(xlEdgeLeft).LineStyle = xlContinuous
        wsOut.Range(wsOut.Cells(1, totalOutCols), wsOut.Cells(outRow + 1, totalOutCols)).Borders(xlEdgeLeft).LineStyle = xlContinuous
    End If
    On Error GoTo Cleanup

    ' Freeze & autofit
    wsOut.Activate: wsOut.Rows("2:2").Select: ActiveWindow.FreezePanes = True
    wsOut.Cells(1, 1).Select: wsOut.UsedRange.Columns.AutoFit

    Dim elapsed As Single
    elapsed = Timer - startTime

    Application.StatusBar = False
    MsgBox "Fuzzy Lookup Complete!" & vbCrLf & vbCrLf & _
           "Rows processed: " & totalLeft & " x " & totalRight & vbCrLf & _
           "Matches found: " & outRow & vbCrLf & _
           "Time: " & Format(elapsed, "0.0") & " seconds" & vbCrLf & _
           "Output: " & outputSheetName, vbInformation, "Fuzzy Lookup"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbCritical, "Fuzzy Lookup"
End Sub
"""

# =====================================================================
# VBA MODULE: MenuSetup (Ribbon tab + menu fallback)
# =====================================================================
MOD_MENU = """
Option Explicit

Private Const MENU_TAG As String = "FuzzyLookupAddin"

Public Sub AddFuzzyLookupMenu()
    RemoveFuzzyLookupMenu

    Dim cbar As CommandBar
    Dim popup As CommandBarPopup
    Dim btn As CommandBarButton

    Set cbar = Application.CommandBars("Worksheet Menu Bar")

    Set popup = cbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    popup.Caption = "Fuzzy Lookup"
    popup.Tag = MENU_TAG

    Set btn = popup.Controls.Add(Type:=msoControlButton)
    btn.Caption = "Open Fuzzy Lookup..."
    btn.OnAction = "ShowFuzzyLookupPane"
    btn.FaceId = 1849
    btn.Style = msoButtonIconAndCaption
    btn.Tag = MENU_TAG

    Set btn = popup.Controls.Add(Type:=msoControlButton)
    btn.Caption = "About"
    btn.OnAction = "ShowAbout"
    btn.FaceId = 487
    btn.Style = msoButtonIconAndCaption
    btn.Tag = MENU_TAG
End Sub

Public Sub RemoveFuzzyLookupMenu()
    Dim ctrl As CommandBarControl
    On Error Resume Next
    For Each ctrl In Application.CommandBars("Worksheet Menu Bar").Controls
        If ctrl.Tag = MENU_TAG Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub

Public Sub ShowFuzzyLookupPane()
    FuzzyTaskPane.Show vbModeless
End Sub

Public Sub ShowAbout()
    MsgBox "Fuzzy Lookup Add-In for Excel" & vbCrLf & _
           "Version 2.1.0 (64-bit compatible)" & vbCrLf & vbCrLf & _
           "A rebuilt replacement for the original Microsoft" & vbCrLf & _
           "Fuzzy Lookup Add-In, fully compatible with" & vbCrLf & _
           "64-bit Excel 2024." & vbCrLf & vbCrLf & _
           "Features:" & vbCrLf & _
           "  - Levenshtein, Jaro-Winkler, Jaccard algorithms" & vbCrLf & _
           "  - Text transformations (find/replace)" & vbCrLf & _
           "  - Configurable similarity threshold" & vbCrLf & _
           "  - Multiple match support" & vbCrLf & vbCrLf & _
           "Worksheet functions:" & vbCrLf & _
           "  =FUZZYMATCH(A1, B1)" & vbCrLf & _
           "  =FUZZYVLOOKUP(A1, range, 1, 2, 0.6)", _
           vbInformation, "About Fuzzy Lookup"
End Sub
"""

# =====================================================================
# VBA: UserForm code — direct design-time control access, no WithEvents
# =====================================================================
FORM_CODE = r"""
Option Explicit

Private Sub UserForm_Initialize()
    Me.StartUpPosition = 0

    On Error Resume Next
    Me.Width = 320
    Me.Height = Application.Height - 60
    Me.Left = Application.Left + Application.Width - Me.Width - 10
    Me.Top = Application.Top + 80
    On Error GoTo 0

    scrThreshold.Min = 0: scrThreshold.Max = 100: scrThreshold.Value = 65
    scrThreshold.SmallChange = 5: scrThreshold.LargeChange = 10
    lblThreshVal.Caption = "0.65"
    txtMaxMatches.Text = "1"
    txtOutputSheet.Text = "Fuzzy_Results"
    chkTrim.Value = True
    chkLower.Value = True
    chkRemovePunct.Value = False
    txtTransforms.Text = "Inc => Incorporated" & vbCrLf & "Corp => Corporation" & vbCrLf & "COM STK =>"

    RefreshTables
End Sub

Private Sub RefreshTables()
    cboLeftTable.Clear
    cboRightTable.Clear

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            cboLeftTable.AddItem lo.Name & "  [Table - " & ws.Name & "]"
            cboRightTable.AddItem lo.Name & "  [Table - " & ws.Name & "]"
        Next lo
    Next ws

    For Each ws In ActiveWorkbook.Worksheets
        If ws.UsedRange.Rows.Count > 1 Then
            cboLeftTable.AddItem ws.Name & "  [Sheet]"
            cboRightTable.AddItem ws.Name & "  [Sheet]"
        End If
    Next ws

    If cboLeftTable.ListCount > 0 Then cboLeftTable.ListIndex = 0
    If cboRightTable.ListCount > 1 Then
        cboRightTable.ListIndex = 1
    ElseIf cboRightTable.ListCount > 0 Then
        cboRightTable.ListIndex = 0
    End If

    ' Explicitly populate columns
    If cboLeftTable.ListIndex >= 0 Then PopulateMatchColumns cboLeftTable, cboLeftMatchCol
    If cboRightTable.ListIndex >= 0 Then PopulateMatchColumns cboRightTable, cboRightMatchCol
End Sub

Private Function ResolveTableRange(ByVal entry As String) As Range
    Dim nm As String
    Dim ws As Worksheet, lo As ListObject

    If InStr(entry, "[Table") > 0 Then
        nm = Trim$(Left$(entry, InStr(entry, "  [") - 1))
        For Each ws In ActiveWorkbook.Worksheets
            For Each lo In ws.ListObjects
                If lo.Name = nm Then
                    Set ResolveTableRange = lo.Range
                    Exit Function
                End If
            Next lo
        Next ws
    ElseIf InStr(entry, "[Sheet]") > 0 Then
        nm = Trim$(Left$(entry, InStr(entry, "  [") - 1))
        Set ws = ActiveWorkbook.Worksheets(nm)
        Set ResolveTableRange = ws.UsedRange
    End If
End Function

Private Sub cboLeftTable_Change()
    PopulateMatchColumns cboLeftTable, cboLeftMatchCol
End Sub

Private Sub cboRightTable_Change()
    PopulateMatchColumns cboRightTable, cboRightMatchCol
End Sub

Private Sub PopulateMatchColumns(cboTable As MSForms.ComboBox, cboCol As MSForms.ComboBox)
    cboCol.Clear
    If cboTable.ListIndex < 0 Then Exit Sub

    Dim rng As Range
    Set rng = ResolveTableRange(cboTable.Value)
    If rng Is Nothing Then Exit Sub

    Dim col As Long
    For col = 1 To rng.Columns.Count
        Dim hdr As String
        hdr = CStr(rng.Cells(1, col).Value)
        If Len(hdr) = 0 Then hdr = "Column " & col
        cboCol.AddItem hdr
    Next col

    If cboCol.ListCount > 0 Then cboCol.ListIndex = 0
End Sub

Private Sub scrThreshold_Change()
    lblThreshVal.Caption = Format(scrThreshold.Value / 100#, "0.00")
End Sub

Private Sub scrThreshold_Scroll()
    scrThreshold_Change
End Sub

Private Sub btnGo_Click()
    If cboLeftTable.ListIndex < 0 Then
        MsgBox "Select a Left Table.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If
    If cboRightTable.ListIndex < 0 Then
        MsgBox "Select a Right Table.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    If cboLeftMatchCol.ListCount = 0 Then PopulateMatchColumns cboLeftTable, cboLeftMatchCol
    If cboRightMatchCol.ListCount = 0 Then PopulateMatchColumns cboRightTable, cboRightMatchCol

    If cboLeftMatchCol.ListIndex < 0 Then
        MsgBox "Select a Left Match Column.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If
    If cboRightMatchCol.ListIndex < 0 Then
        MsgBox "Select a Right Match Column.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    Dim leftRange As Range, rightRange As Range
    Set leftRange = ResolveTableRange(cboLeftTable.Value)
    Set rightRange = ResolveTableRange(cboRightTable.Value)

    If leftRange Is Nothing Or rightRange Is Nothing Then
        MsgBox "Could not resolve table ranges.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    Dim threshold As Double: threshold = scrThreshold.Value / 100#
    Dim maxMatches As Long
    If IsNumeric(txtMaxMatches.Text) Then maxMatches = CLng(txtMaxMatches.Text) Else maxMatches = 1
    If maxMatches < 1 Then maxMatches = 1
    Dim outputName As String: outputName = Trim$(txtOutputSheet.Text)
    If Len(outputName) = 0 Then outputName = "Fuzzy_Results"

    Dim rules() As TransformRule
    Dim ruleCount As Long
    ruleCount = 0
    ReDim rules(0 To 99)
    Dim transText As String: transText = txtTransforms.Text
    If Len(Trim$(transText)) > 0 Then
        Dim lines() As String: lines = Split(transText, vbCrLf)
        Dim ln As Long
        For ln = LBound(lines) To UBound(lines)
            Dim parts() As String
            Dim line As String: line = Trim$(lines(ln))
            If Len(line) > 0 And InStr(line, "=>") > 0 Then
                parts = Split(line, "=>", 2)
                rules(ruleCount).FromText = Trim$(parts(0))
                If UBound(parts) >= 1 Then rules(ruleCount).ToText = Trim$(parts(1)) Else rules(ruleCount).ToText = ""
                ruleCount = ruleCount + 1
            End If
        Next ln
    End If

    ExecuteFuzzyLookup leftRange.Worksheet, leftRange.Address, _
                       cboLeftMatchCol.ListIndex + 1, _
                       rightRange.Worksheet, rightRange.Address, _
                       cboRightMatchCol.ListIndex + 1, _
                       threshold, maxMatches, _
                       chkTrim.Value, chkLower.Value, chkRemovePunct.Value, _
                       outputName, rules, ruleCount
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub
"""

# =====================================================================
# VBA: ThisWorkbook events
# =====================================================================
THISWORKBOOK_CODE = """
Private Sub Workbook_Open()
    AddFuzzyLookupMenu
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    RemoveFuzzyLookupMenu
End Sub
"""


def add_designer_label(designer, name, caption, left, top, width, height=15,
                       font_size=9, bold=False, back_color=None, fore_color=None,
                       text_align=None):
    ctrl = designer.Controls.Add("Forms.Label.1", name, True)
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Font.Size = font_size
    if bold:
        ctrl.Font.Bold = True
    if back_color is not None:
        ctrl.BackColor = back_color
    if fore_color is not None:
        ctrl.ForeColor = fore_color
    if text_align is not None:
        ctrl.TextAlign = text_align
    return ctrl


def add_designer_combobox(designer, name, left, top, width, height=20):
    ctrl = designer.Controls.Add("Forms.ComboBox.1", name, True)
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Style = 2  # fmStyleDropDownList
    ctrl.Font.Size = 9
    return ctrl


def add_designer_textbox(designer, name, left, top, width, height=20, text=""):
    ctrl = designer.Controls.Add("Forms.TextBox.1", name, True)
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = width
    ctrl.Height = height
    ctrl.Text = text
    ctrl.Font.Size = 9
    return ctrl


def add_designer_checkbox(designer, name, caption, left, top, value=False):
    ctrl = designer.Controls.Add("Forms.CheckBox.1", name, True)
    ctrl.Caption = caption
    ctrl.Left = left
    ctrl.Top = top
    ctrl.Width = 280
    ctrl.Height = 18
    ctrl.Value = value
    ctrl.Font.Size = 9
    return ctrl


def add_designer_section(designer, name, caption, y):
    return add_designer_label(designer, name, caption, 4, y, 304, 18,
                              font_size=9, bold=True,
                              back_color=0xC47244,  # RGB(68,114,196) BGR
                              fore_color=0xFFFFFF)


def build_form_controls(designer):
    """Place all controls on the form at build/design time."""
    y = 8

    # LEFT TABLE
    add_designer_section(designer, "lblLeftHdr", "Left Table (Lookup Values)", y)
    y += 22
    add_designer_label(designer, "lblLeftTbl", "Table:", 8, y + 3, 42)
    add_designer_combobox(designer, "cboLeftTable", 54, y, 250)
    y += 28
    add_designer_label(designer, "lblLeftCol", "Match Column:", 8, y + 3, 80)
    add_designer_combobox(designer, "cboLeftMatchCol", 92, y, 212)
    y += 34

    # RIGHT TABLE
    add_designer_section(designer, "lblRightHdr", "Right Table (Match Against)", y)
    y += 22
    add_designer_label(designer, "lblRightTbl", "Table:", 8, y + 3, 42)
    add_designer_combobox(designer, "cboRightTable", 54, y, 250)
    y += 28
    add_designer_label(designer, "lblRightCol", "Match Column:", 8, y + 3, 80)
    add_designer_combobox(designer, "cboRightMatchCol", 92, y, 212)
    y += 34

    # SIMILARITY
    add_designer_section(designer, "lblSimHdr", "Similarity Threshold", y)
    y += 22
    scr = designer.Controls.Add("Forms.ScrollBar.1", "scrThreshold", True)
    scr.Left = 8; scr.Top = y; scr.Width = 220; scr.Height = 20
    scr.Min = 0; scr.Max = 100; scr.Value = 65
    scr.SmallChange = 5; scr.LargeChange = 10
    scr.Orientation = 1  # fmOrientationHorizontal
    add_designer_label(designer, "lblThreshVal", "0.65", 234, y + 2, 40,
                       font_size=11, bold=True)
    y += 28
    add_designer_label(designer, "lblLow", "Low (more matches)", 8, y, 130, font_size=7)
    add_designer_label(designer, "lblHigh", "High (fewer matches)", 168, y, 130,
                       font_size=7, text_align=3)
    y += 22

    # MAX MATCHES
    add_designer_label(designer, "lblMaxM", "Max Matches:", 8, y + 3, 78)
    add_designer_textbox(designer, "txtMaxMatches", 90, y, 40, text="1")
    add_designer_label(designer, "lblMaxMHint", "per row", 134, y + 3, 50)
    y += 30

    # OUTPUT
    add_designer_label(designer, "lblOutSheet", "Output Sheet:", 8, y + 3, 78)
    add_designer_textbox(designer, "txtOutputSheet", 90, y, 140, text="Fuzzy_Results")
    y += 34

    # TEXT CLEANING
    add_designer_section(designer, "lblTransHdr", "Text Cleaning", y)
    y += 22
    add_designer_checkbox(designer, "chkTrim", "Trim && collapse spaces", 8, y, True)
    y += 20
    add_designer_checkbox(designer, "chkLower", "Convert to lowercase", 8, y, True)
    y += 20
    add_designer_checkbox(designer, "chkRemovePunct", "Remove punctuation", 8, y, False)
    y += 28

    # CUSTOM REPLACEMENTS
    add_designer_section(designer, "lblCustHdr", "Custom Replacements (optional)", y)
    y += 20
    add_designer_label(designer, "lblCustHelp", "One per line:  find => replace", 8, y, 280,
                       font_size=7)
    y += 16
    txt = add_designer_textbox(designer, "txtTransforms", 8, y, 296, height=60,
                               text="Inc => Incorporated\r\nCorp => Corporation\r\nCOM STK =>")
    txt.MultiLine = True
    txt.ScrollBars = 2  # fmScrollBarsVertical
    txt.Font.Name = "Consolas"
    txt.Font.Size = 9
    y += 68

    # BUTTONS
    btn_go = designer.Controls.Add("Forms.CommandButton.1", "btnGo", True)
    btn_go.Caption = "Go!"
    btn_go.Left = 8; btn_go.Top = y; btn_go.Width = 200; btn_go.Height = 36
    btn_go.Font.Size = 13; btn_go.Font.Bold = True
    btn_go.BackColor = 0xC47244; btn_go.ForeColor = 0xFFFFFF

    btn_close = designer.Controls.Add("Forms.CommandButton.1", "btnClose", True)
    btn_close.Caption = "Close"
    btn_close.Left = 214; btn_close.Top = y; btn_close.Width = 90; btn_close.Height = 36
    btn_close.Font.Size = 11


def build_addin():
    print("=" * 60)
    print("Building Fuzzy Lookup Add-In v2.3")
    print("64-bit Excel 2024 Compatible")
    print("Designer controls + WithEvents for events")
    print("=" * 60)

    if os.path.exists(OUTPUT_PATH):
        try:
            os.remove(OUTPUT_PATH)
        except Exception as e:
            print(f"Warning: {e}")

    excel = None
    wb = None
    try:
        print("\nStarting Excel...")
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        print("Creating workbook...")
        wb = excel.Workbooks.Add()
        vbp = wb.VBProject

        # Add standard modules
        modules = {
            "FuzzyAlgorithms": MOD_ALGORITHMS,
            "FuzzyEngine": MOD_ENGINE,
            "MenuSetup": MOD_MENU,
        }

        for name, code in modules.items():
            print(f"  Adding module: {name}")
            mod = vbp.VBComponents.Add(1)  # vbext_ct_StdModule
            mod.Name = name
            mod.CodeModule.AddFromString(code.strip())

        # Create UserForm and add controls via Designer (design-time)
        print("  Creating Task Pane (UserForm)...")
        frm = vbp.VBComponents.Add(3)  # vbext_ct_MSForm
        frm.Name = "FuzzyTaskPane"
        frm.Properties.Item("Caption").Value = "Fuzzy Lookup"
        frm.Properties.Item("Width").Value = 320
        frm.Properties.Item("Height").Value = 620
        frm.Properties.Item("BackColor").Value = 0xFFFFFF

        print("  Adding controls via Designer (design-time)...")
        build_form_controls(frm.Designer)

        # Add form code with WithEvents to wire events to design-time controls
        print("  Adding form code (WithEvents event wiring)...")
        cm = frm.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        cm.AddFromString(FORM_CODE.strip())

        # Update ThisWorkbook
        print("  Adding workbook events...")
        tw = vbp.VBComponents.Item("ThisWorkbook")
        twcm = tw.CodeModule
        if twcm.CountOfLines > 0:
            twcm.DeleteLines(1, twcm.CountOfLines)
        twcm.AddFromString(THISWORKBOOK_CODE.strip())

        # Save as .xlam
        print(f"\nSaving: {OUTPUT_PATH}")
        wb.SaveAs(OUTPUT_PATH, FileFormat=55)
        print("Save successful!")

        wb.Close(False)
        wb = None

        print("\n" + "=" * 60)
        print("BUILD COMPLETE!")
        print("=" * 60)
        print(f"\nFile: {OUTPUT_PATH}")
        print(f"Size: {os.path.getsize(OUTPUT_PATH) / 1024:.0f} KB")
        print("\nInstall:")
        print("  1. Open Excel > File > Options > Add-ins")
        print("  2. Manage: Excel Add-ins > Go")
        print("  3. Browse > select FuzzyLookup.xlam > OK")
        print("\nUsage:")
        print("  - Menu bar: Fuzzy Lookup > Open Fuzzy Lookup...")
        print("  - Task pane docks to right side of Excel")
        print("  - Select tables, columns, threshold, and click Go!")
        print("  - Cell formulas: =FUZZYMATCH(A1,B1)  =FUZZYVLOOKUP(...)")

    except Exception as e:
        print(f"\nERROR: {e}")
        raise
    finally:
        if wb:
            try: wb.Close(False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass


if __name__ == "__main__":
    build_addin()
