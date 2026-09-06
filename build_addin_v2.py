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
    Dim i As Long
    Dim result As String
    result = text
    For i = 0 To ruleCount - 1
        If Len(rules(i).FromText) > 0 Then
            result = Replace(result, rules(i).FromText, rules(i).ToText, , , vbTextCompare)
        End If
    Next i
    ApplyTransformations = result
End Function

Public Function CleanText(ByVal text As String, _
                          trimSpaces As Boolean, _
                          toLower As Boolean, _
                          removePunct As Boolean) As String
    Dim result As String
    result = text

    If toLower Then result = LCase$(result)

    If removePunct Then
        Dim i As Long, ch As String, cleaned As String
        cleaned = ""
        For i = 1 To Len(result)
            ch = Mid$(result, i, 1)
            Select Case Asc(ch)
                Case 32, 48 To 57, 65 To 90, 97 To 122
                    cleaned = cleaned & ch
            End Select
        Next i
        result = cleaned
    End If

    If trimSpaces Then
        result = Trim$(result)
        Do While InStr(result, "  ") > 0
            result = Replace(result, "  ", " ")
        Loop
    End If

    CleanText = result
End Function

Public Sub ExecuteFuzzyLookup(leftWs As Worksheet, leftAddr As String, leftMatchCol As Long, _
                               rightWs As Worksheet, rightAddr As String, rightMatchCol As Long, _
                               threshold As Double, maxMatches As Long, _
                               doTrim As Boolean, doLower As Boolean, doRemovePunct As Boolean, _
                               outputSheetName As String, _
                               rules() As TransformRule, ruleCount As Long)

    Dim leftRange As Range, rightRange As Range
    Dim wsOut As Worksheet
    Dim leftRow As Long, rightRow As Long
    Dim leftVal As String, rightVal As String
    Dim score As Double
    Dim outRow As Long, c As Long
    Dim leftCols As Long, rightCols As Long
    Dim totalLeft As Long, totalRight As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo Cleanup

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

    ' Create/clear output sheet
    On Error Resume Next
    Set wsOut = ActiveWorkbook.Sheets(outputSheetName)
    On Error GoTo Cleanup
    If wsOut Is Nothing Then
        Set wsOut = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
        wsOut.Name = outputSheetName
    Else
        wsOut.Cells.Clear
    End If

    ' Headers
    outRow = 1: c = 1
    Dim colIdx As Long
    For colIdx = 1 To leftCols
        wsOut.Cells(outRow, c).Value = leftRange.Cells(1, colIdx).Value
        wsOut.Cells(outRow, c).Interior.Color = RGB(68, 114, 196)
        wsOut.Cells(outRow, c).Font.Color = RGB(255, 255, 255)
        wsOut.Cells(outRow, c).Font.Bold = True
        c = c + 1
    Next colIdx
    For colIdx = 1 To rightCols
        wsOut.Cells(outRow, c).Value = rightRange.Cells(1, colIdx).Value
        wsOut.Cells(outRow, c).Interior.Color = RGB(47, 117, 181)
        wsOut.Cells(outRow, c).Font.Color = RGB(255, 255, 255)
        wsOut.Cells(outRow, c).Font.Bold = True
        c = c + 1
    Next colIdx
    wsOut.Cells(outRow, c).Value = "Similarity"
    wsOut.Cells(outRow, c).Interior.Color = RGB(0, 176, 80)
    wsOut.Cells(outRow, c).Font.Color = RGB(255, 255, 255)
    wsOut.Cells(outRow, c).Font.Bold = True

    outRow = 2

    ' Pre-cache and pre-clean right table match column
    Dim rightCache() As String
    ReDim rightCache(1 To totalRight)
    For rightRow = 1 To totalRight
        rightVal = CStr(rightRange.Cells(rightRow + 1, rightMatchCol).Value)
        rightVal = CleanText(rightVal, doTrim, doLower, doRemovePunct)
        If ruleCount > 0 Then rightVal = ApplyTransformations(rightVal, rules, ruleCount)
        rightCache(rightRow) = rightVal
    Next rightRow

    ' Main matching loop
    Dim matchScores() As Double, matchRows() As Long
    ReDim matchScores(1 To totalRight)
    ReDim matchRows(1 To totalRight)

    For leftRow = 1 To totalLeft
        If leftRow Mod 5 = 0 Or leftRow = totalLeft Then
            Application.StatusBar = "Fuzzy Lookup: " & leftRow & " / " & totalLeft & _
                                    " (" & Format(CDbl(leftRow) / CDbl(totalLeft) * 100, "0") & "%)"
            DoEvents
        End If

        leftVal = CStr(leftRange.Cells(leftRow + 1, leftMatchCol).Value)
        leftVal = CleanText(leftVal, doTrim, doLower, doRemovePunct)
        If ruleCount > 0 Then leftVal = ApplyTransformations(leftVal, rules, ruleCount)

        Dim matchCount As Long
        matchCount = 0

        For rightRow = 1 To totalRight
            score = CombinedSimilarity(leftVal, rightCache(rightRow))
            If score >= threshold Then
                matchCount = matchCount + 1
                matchScores(matchCount) = score
                matchRows(matchCount) = rightRow
            End If
        Next rightRow

        ' Sort matches descending by score (insertion sort)
        If matchCount > 1 Then
            Dim ii As Long, jj As Long, tmpS As Double, tmpR As Long
            For ii = 2 To matchCount
                tmpS = matchScores(ii): tmpR = matchRows(ii)
                jj = ii - 1
                Do While jj >= 1
                    If matchScores(jj) < tmpS Then
                        matchScores(jj + 1) = matchScores(jj)
                        matchRows(jj + 1) = matchRows(jj)
                        jj = jj - 1
                    Else
                        Exit Do
                    End If
                Loop
                matchScores(jj + 1) = tmpS: matchRows(jj + 1) = tmpR
            Next ii
        End If

        Dim outputN As Long
        outputN = matchCount: If outputN > maxMatches Then outputN = maxMatches

        For ii = 1 To outputN
            c = 1
            Dim rowColor As Long
            If (outRow Mod 2) = 0 Then rowColor = RGB(242, 246, 252) Else rowColor = RGB(255, 255, 255)

            For colIdx = 1 To leftCols
                wsOut.Cells(outRow, c).Value = leftRange.Cells(leftRow + 1, colIdx).Value
                wsOut.Cells(outRow, c).Interior.Color = rowColor
                c = c + 1
            Next colIdx
            For colIdx = 1 To rightCols
                wsOut.Cells(outRow, c).Value = rightRange.Cells(matchRows(ii) + 1, colIdx).Value
                wsOut.Cells(outRow, c).Interior.Color = rowColor
                c = c + 1
            Next colIdx

            wsOut.Cells(outRow, c).Value = Round(matchScores(ii), 4)
            wsOut.Cells(outRow, c).NumberFormat = "0.00%"
            wsOut.Cells(outRow, c).Interior.Color = rowColor
            If matchScores(ii) >= 0.9 Then
                wsOut.Cells(outRow, c).Font.Color = RGB(0, 128, 0)
            ElseIf matchScores(ii) >= 0.7 Then
                wsOut.Cells(outRow, c).Font.Color = RGB(0, 0, 0)
            Else
                wsOut.Cells(outRow, c).Font.Color = RGB(200, 100, 0)
            End If
            wsOut.Cells(outRow, c).Font.Bold = True
            outRow = outRow + 1
        Next ii
    Next leftRow

    ' Column borders
    Dim totalCols As Long
    totalCols = leftCols + rightCols + 1
    With wsOut.Range(wsOut.Cells(1, leftCols + 1), wsOut.Cells(outRow - 1, leftCols + 1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With
    With wsOut.Range(wsOut.Cells(1, leftCols + rightCols + 1), wsOut.Cells(outRow - 1, leftCols + rightCols + 1)).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(100, 100, 100)
    End With

    ' Freeze header
    wsOut.Activate
    wsOut.Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    wsOut.Cells(1, 1).Select

    wsOut.UsedRange.Columns.AutoFit

    Application.StatusBar = False
    MsgBox "Fuzzy Lookup Complete!" & vbCrLf & vbCrLf & _
           "Rows processed: " & totalLeft & vbCrLf & _
           "Matches found: " & (outRow - 2) & vbCrLf & _
           "Output: " & outputSheetName, vbInformation, "Fuzzy Lookup"

Cleanup:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Description, vbCritical, "Fuzzy Lookup"
    End If
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
# VBA: UserForm code — WithEvents for reliable event handling
# All controls created dynamically, events wired via WithEvents vars
# =====================================================================
FORM_CODE = r"""
Option Explicit

' WithEvents declarations for controls that need event handling
Private WithEvents mBtnGo As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton
Private WithEvents mCboLeftTable As MSForms.ComboBox
Private WithEvents mCboRightTable As MSForms.ComboBox
Private WithEvents mScrThreshold As MSForms.ScrollBar

Private Sub UserForm_Initialize()
    Me.Caption = "Fuzzy Lookup"
    Me.StartUpPosition = 0
    Me.BackColor = RGB(255, 255, 255)
    PositionAsTaskPane
    BuildControls
    RefreshTables
End Sub

Private Sub PositionAsTaskPane()
    On Error Resume Next
    Me.Width = 320
    Me.Height = Application.Height - 60
    Me.Left = Application.Left + Application.Width - Me.Width - 10
    Me.Top = Application.Top + 80
    On Error GoTo 0
End Sub

Private Sub BuildControls()
    Dim y As Single
    y = 8

    ' ---- LEFT TABLE SECTION ----
    AddSectionHeader "lblLeftHdr", "Left Table (Lookup Values)", y
    y = y + 22

    AddLabel "lblLeftTbl", "Table:", 8, y + 3, 42
    Set mCboLeftTable = Me.Controls.Add("Forms.ComboBox.1", "cboLeftTable")
    mCboLeftTable.Left = 54: mCboLeftTable.Top = y: mCboLeftTable.Width = 250: mCboLeftTable.Height = 20
    mCboLeftTable.Style = fmStyleDropDownList: mCboLeftTable.Font.Size = 9
    y = y + 28

    AddLabel "lblLeftCol", "Match Column:", 8, y + 3, 80
    AddComboBox "cboLeftMatchCol", 92, y, 212
    y = y + 34

    ' ---- RIGHT TABLE SECTION ----
    AddSectionHeader "lblRightHdr", "Right Table (Match Against)", y
    y = y + 22

    AddLabel "lblRightTbl", "Table:", 8, y + 3, 42
    Set mCboRightTable = Me.Controls.Add("Forms.ComboBox.1", "cboRightTable")
    mCboRightTable.Left = 54: mCboRightTable.Top = y: mCboRightTable.Width = 250: mCboRightTable.Height = 20
    mCboRightTable.Style = fmStyleDropDownList: mCboRightTable.Font.Size = 9
    y = y + 28

    AddLabel "lblRightCol", "Match Column:", 8, y + 3, 80
    AddComboBox "cboRightMatchCol", 92, y, 212
    y = y + 34

    ' ---- SIMILARITY SECTION ----
    AddSectionHeader "lblSimHdr", "Similarity Threshold", y
    y = y + 22

    Set mScrThreshold = Me.Controls.Add("Forms.ScrollBar.1", "scrThreshold")
    mScrThreshold.Left = 8: mScrThreshold.Top = y: mScrThreshold.Width = 220: mScrThreshold.Height = 20
    mScrThreshold.Min = 0: mScrThreshold.Max = 100: mScrThreshold.Value = 65
    mScrThreshold.SmallChange = 5: mScrThreshold.LargeChange = 10
    mScrThreshold.Orientation = fmOrientationHorizontal

    AddLabel "lblThreshVal", "0.65", 234, y + 2, 40
    Me.Controls("lblThreshVal").Font.Bold = True
    Me.Controls("lblThreshVal").Font.Size = 11
    y = y + 28

    AddLabel "lblLow", "Low (more matches)", 8, y, 130
    Me.Controls("lblLow").Font.Size = 7
    Dim lblHigh As MSForms.Label
    Set lblHigh = Me.Controls.Add("Forms.Label.1", "lblHigh")
    lblHigh.Caption = "High (fewer matches)"
    lblHigh.Left = 168: lblHigh.Top = y: lblHigh.Width = 130: lblHigh.Height = 14
    lblHigh.Font.Size = 7: lblHigh.TextAlign = fmTextAlignRight
    y = y + 22

    ' ---- MAX MATCHES ----
    AddLabel "lblMaxM", "Max Matches:", 8, y + 3, 78
    AddTextBox "txtMaxMatches", 90, y, 40, "1"
    AddLabel "lblMaxMHint", "per row", 134, y + 3, 50
    y = y + 30

    ' ---- OUTPUT SECTION ----
    AddLabel "lblOutSheet", "Output Sheet:", 8, y + 3, 78
    AddTextBox "txtOutputSheet", 90, y, 140, "Fuzzy_Results"
    y = y + 34

    ' ---- TEXT CLEANING SECTION ----
    AddSectionHeader "lblTransHdr", "Text Cleaning", y
    y = y + 22

    AddCheckBox "chkTrim", "Trim & collapse spaces", 8, y, True
    y = y + 20
    AddCheckBox "chkLower", "Convert to lowercase", 8, y, True
    y = y + 20
    AddCheckBox "chkRemovePunct", "Remove punctuation", 8, y, False
    y = y + 28

    ' ---- CUSTOM TRANSFORMATIONS ----
    AddSectionHeader "lblCustHdr", "Custom Replacements (optional)", y
    y = y + 20
    AddLabel "lblCustHelp", "One per line:  find => replace", 8, y, 280
    Me.Controls("lblCustHelp").Font.Size = 7
    y = y + 16

    Dim txtTrans As MSForms.TextBox
    Set txtTrans = Me.Controls.Add("Forms.TextBox.1", "txtTransforms")
    txtTrans.Left = 8: txtTrans.Top = y: txtTrans.Width = 296: txtTrans.Height = 60
    txtTrans.MultiLine = True: txtTrans.ScrollBars = fmScrollBarsVertical
    txtTrans.Text = "Inc => Incorporated" & vbCrLf & "Corp => Corporation" & vbCrLf & "COM STK =>"
    txtTrans.Font.Name = "Consolas": txtTrans.Font.Size = 9
    y = y + 68

    ' ---- BUTTONS ----
    Set mBtnGo = Me.Controls.Add("Forms.CommandButton.1", "btnGo")
    mBtnGo.Caption = "Go!"
    mBtnGo.Left = 8: mBtnGo.Top = y: mBtnGo.Width = 200: mBtnGo.Height = 36
    mBtnGo.Font.Size = 13: mBtnGo.Font.Bold = True
    mBtnGo.BackColor = RGB(68, 114, 196): mBtnGo.ForeColor = RGB(255, 255, 255)

    Set mBtnClose = Me.Controls.Add("Forms.CommandButton.1", "btnClose")
    mBtnClose.Caption = "Close"
    mBtnClose.Left = 214: mBtnClose.Top = y: mBtnClose.Width = 90: mBtnClose.Height = 36
    mBtnClose.Font.Size = 11
End Sub

' ---- Helper subs for non-event controls ----
Private Sub AddSectionHeader(nm As String, cap As String, y As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", nm)
    lbl.Caption = cap
    lbl.Left = 4: lbl.Top = y: lbl.Width = 304: lbl.Height = 18
    lbl.Font.Bold = True: lbl.Font.Size = 9
    lbl.BackColor = RGB(68, 114, 196): lbl.ForeColor = RGB(255, 255, 255)
End Sub

Private Sub AddLabel(nm As String, cap As String, l As Single, t As Single, w As Single)
    Dim lbl As MSForms.Label
    Set lbl = Me.Controls.Add("Forms.Label.1", nm)
    lbl.Caption = cap
    lbl.Left = l: lbl.Top = t: lbl.Width = w: lbl.Height = 15: lbl.Font.Size = 9
End Sub

Private Sub AddComboBox(nm As String, l As Single, t As Single, w As Single)
    Dim cbo As MSForms.ComboBox
    Set cbo = Me.Controls.Add("Forms.ComboBox.1", nm)
    cbo.Left = l: cbo.Top = t: cbo.Width = w: cbo.Height = 20
    cbo.Style = fmStyleDropDownList: cbo.Font.Size = 9
End Sub

Private Sub AddTextBox(nm As String, l As Single, t As Single, w As Single, def As String)
    Dim txt As MSForms.TextBox
    Set txt = Me.Controls.Add("Forms.TextBox.1", nm)
    txt.Left = l: txt.Top = t: txt.Width = w: txt.Height = 20
    txt.Text = def: txt.Font.Size = 9
End Sub

Private Sub AddCheckBox(nm As String, cap As String, l As Single, t As Single, defaultVal As Boolean)
    Dim chk As MSForms.CheckBox
    Set chk = Me.Controls.Add("Forms.CheckBox.1", nm)
    chk.Caption = cap
    chk.Left = l: chk.Top = t: chk.Width = 280: chk.Height = 18
    chk.Value = defaultVal: chk.Font.Size = 9
End Sub

' ---- Populate table/range list ----
Private Sub RefreshTables()
    mCboLeftTable.Clear
    mCboRightTable.Clear

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            mCboLeftTable.AddItem lo.Name & "  [Table - " & ws.Name & "]"
            mCboRightTable.AddItem lo.Name & "  [Table - " & ws.Name & "]"
        Next lo
    Next ws

    For Each ws In ActiveWorkbook.Worksheets
        If ws.UsedRange.Rows.Count > 1 Then
            mCboLeftTable.AddItem ws.Name & "  [Sheet]"
            mCboRightTable.AddItem ws.Name & "  [Sheet]"
        End If
    Next ws

    If mCboLeftTable.ListCount > 0 Then mCboLeftTable.ListIndex = 0
    If mCboRightTable.ListCount > 1 Then
        mCboRightTable.ListIndex = 1
    ElseIf mCboRightTable.ListCount > 0 Then
        mCboRightTable.ListIndex = 0
    End If
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

' ---- WithEvents handlers ----
Private Sub mCboLeftTable_Change()
    PopulateMatchColumns mCboLeftTable, Me.Controls("cboLeftMatchCol")
End Sub

Private Sub mCboRightTable_Change()
    PopulateMatchColumns mCboRightTable, Me.Controls("cboRightMatchCol")
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

Private Sub mScrThreshold_Change()
    Me.Controls("lblThreshVal").Caption = Format(mScrThreshold.Value / 100#, "0.00")
End Sub

Private Sub mScrThreshold_Scroll()
    mScrThreshold_Change
End Sub

Private Sub mBtnGo_Click()
    Dim cboLT As MSForms.ComboBox: Set cboLT = mCboLeftTable
    Dim cboRT As MSForms.ComboBox: Set cboRT = mCboRightTable
    Dim cboLC As MSForms.ComboBox: Set cboLC = Me.Controls("cboLeftMatchCol")
    Dim cboRC As MSForms.ComboBox: Set cboRC = Me.Controls("cboRightMatchCol")

    If cboLT.ListIndex < 0 Then
        MsgBox "Select a Left Table.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If
    If cboRT.ListIndex < 0 Then
        MsgBox "Select a Right Table.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    ' Auto-populate columns if empty
    If cboLC.ListCount = 0 Then PopulateMatchColumns cboLT, cboLC
    If cboRC.ListCount = 0 Then PopulateMatchColumns cboRT, cboRC

    If cboLC.ListIndex < 0 Then
        MsgBox "Select a Left Match Column.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If
    If cboRC.ListIndex < 0 Then
        MsgBox "Select a Right Match Column.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    Dim leftRange As Range, rightRange As Range
    Set leftRange = ResolveTableRange(cboLT.Value)
    Set rightRange = ResolveTableRange(cboRT.Value)

    If leftRange Is Nothing Or rightRange Is Nothing Then
        MsgBox "Could not resolve table ranges.", vbExclamation, "Fuzzy Lookup": Exit Sub
    End If

    Dim threshold As Double
    threshold = mScrThreshold.Value / 100#

    Dim maxMatches As Long
    If IsNumeric(Me.Controls("txtMaxMatches").Text) Then
        maxMatches = CLng(Me.Controls("txtMaxMatches").Text)
    Else
        maxMatches = 1
    End If
    If maxMatches < 1 Then maxMatches = 1

    Dim outputName As String
    outputName = Trim$(Me.Controls("txtOutputSheet").Text)
    If Len(outputName) = 0 Then outputName = "Fuzzy_Results"

    ' Parse transformations
    Dim rules() As TransformRule
    Dim ruleCount As Long
    ruleCount = 0
    ReDim rules(0 To 99)

    Dim transText As String
    transText = Me.Controls("txtTransforms").Text
    If Len(Trim$(transText)) > 0 Then
        Dim lines() As String
        lines = Split(transText, vbCrLf)
        Dim ln As Long
        For ln = LBound(lines) To UBound(lines)
            Dim parts() As String
            Dim line As String
            line = Trim$(lines(ln))
            If Len(line) > 0 And InStr(line, "=>") > 0 Then
                parts = Split(line, "=>", 2)
                rules(ruleCount).FromText = Trim$(parts(0))
                If UBound(parts) >= 1 Then rules(ruleCount).ToText = Trim$(parts(1)) Else rules(ruleCount).ToText = ""
                ruleCount = ruleCount + 1
            End If
        Next ln
    End If

    Dim doTrim As Boolean, doLower As Boolean, doRemovePunct As Boolean
    doTrim = Me.Controls("chkTrim").Value
    doLower = Me.Controls("chkLower").Value
    doRemovePunct = Me.Controls("chkRemovePunct").Value

    ExecuteFuzzyLookup leftRange.Worksheet, leftRange.Address, _
                       cboLC.ListIndex + 1, _
                       rightRange.Worksheet, rightRange.Address, _
                       cboRC.ListIndex + 1, _
                       threshold, maxMatches, _
                       doTrim, doLower, doRemovePunct, _
                       outputName, rules, ruleCount
End Sub

Private Sub mBtnClose_Click()
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


def build_addin():
    print("=" * 60)
    print("Building Fuzzy Lookup Add-In v2.2")
    print("64-bit Excel 2024 Compatible")
    print("Fix: WithEvents for reliable event handling")
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

        # Create blank UserForm (controls added dynamically via VBA with WithEvents)
        print("  Creating Task Pane (UserForm)...")
        frm = vbp.VBComponents.Add(3)  # vbext_ct_MSForm
        frm.Name = "FuzzyTaskPane"
        frm.Properties.Item("Caption").Value = "Fuzzy Lookup"
        frm.Properties.Item("Width").Value = 320
        frm.Properties.Item("Height").Value = 620
        frm.Properties.Item("BackColor").Value = 0xFFFFFF

        # Add WithEvents form code
        print("  Adding form code with WithEvents event handling...")
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
