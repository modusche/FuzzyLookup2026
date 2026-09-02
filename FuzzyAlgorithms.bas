Attribute VB_Name = "FuzzyAlgorithms"
Option Explicit

' ============================================================
' Fuzzy Matching Algorithms Module
' Compatible with 64-bit Excel 2024
' ============================================================

Public Function LevenshteinDistance(ByVal s1 As String, ByVal s2 As String) As Long
    Dim len1 As Long, len2 As Long
    Dim matrix() As Long
    Dim i As Long, j As Long
    Dim cost As Long
    Dim above As Long, left As Long, diag As Long

    s1 = LCase$(s1)
    s2 = LCase$(s2)
    len1 = Len(s1)
    len2 = Len(s2)

    If len1 = 0 Then
        LevenshteinDistance = len2
        Exit Function
    End If
    If len2 = 0 Then
        LevenshteinDistance = len1
        Exit Function
    End If

    ReDim matrix(0 To len1, 0 To len2)

    For i = 0 To len1
        matrix(i, 0) = i
    Next i
    For j = 0 To len2
        matrix(0, j) = j
    Next j

    For i = 1 To len1
        For j = 1 To len2
            If Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                cost = 0
            Else
                cost = 1
            End If
            above = matrix(i - 1, j) + 1
            left = matrix(i, j - 1) + 1
            diag = matrix(i - 1, j - 1) + cost

            matrix(i, j) = above
            If left < matrix(i, j) Then matrix(i, j) = left
            If diag < matrix(i, j) Then matrix(i, j) = diag
        Next j
    Next i

    LevenshteinDistance = matrix(len1, len2)
End Function

Public Function LevenshteinSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim maxLen As Long
    maxLen = Len(s1)
    If Len(s2) > maxLen Then maxLen = Len(s2)
    If maxLen = 0 Then
        LevenshteinSimilarity = 1#
        Exit Function
    End If
    LevenshteinSimilarity = 1# - (CDbl(LevenshteinDistance(s1, s2)) / CDbl(maxLen))
End Function

Public Function JaroSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    Dim len1 As Long, len2 As Long
    Dim matchDist As Long
    Dim matches As Long, transpositions As Long
    Dim s1Matches() As Boolean, s2Matches() As Boolean
    Dim i As Long, j As Long, k As Long
    Dim startJ As Long, endJ As Long

    s1 = LCase$(s1)
    s2 = LCase$(s2)
    len1 = Len(s1)
    len2 = Len(s2)

    If len1 = 0 And len2 = 0 Then
        JaroSimilarity = 1#
        Exit Function
    End If
    If len1 = 0 Or len2 = 0 Then
        JaroSimilarity = 0#
        Exit Function
    End If

    If len1 > len2 Then
        matchDist = CLng(len1 / 2) - 1
    Else
        matchDist = CLng(len2 / 2) - 1
    End If
    If matchDist < 0 Then matchDist = 0

    ReDim s1Matches(1 To len1)
    ReDim s2Matches(1 To len2)

    matches = 0
    transpositions = 0

    For i = 1 To len1
        startJ = i - matchDist
        If startJ < 1 Then startJ = 1
        endJ = i + matchDist
        If endJ > len2 Then endJ = len2

        For j = startJ To endJ
            If Not s2Matches(j) And Mid$(s1, i, 1) = Mid$(s2, j, 1) Then
                s1Matches(i) = True
                s2Matches(j) = True
                matches = matches + 1
                Exit For
            End If
        Next j
    Next i

    If matches = 0 Then
        JaroSimilarity = 0#
        Exit Function
    End If

    k = 1
    For i = 1 To len1
        If s1Matches(i) Then
            Do While Not s2Matches(k)
                k = k + 1
            Loop
            If Mid$(s1, i, 1) <> Mid$(s2, k, 1) Then
                transpositions = transpositions + 1
            End If
            k = k + 1
        End If
    Next i

    JaroSimilarity = (CDbl(matches) / CDbl(len1) + _
                      CDbl(matches) / CDbl(len2) + _
                      (CDbl(matches) - CDbl(transpositions) / 2#) / CDbl(matches)) / 3#
End Function

Public Function JaroWinklerSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                       Optional ByVal scalingFactor As Double = 0.1) As Double
    Dim jaro As Double
    Dim prefixLen As Long
    Dim maxPrefix As Long
    Dim i As Long

    jaro = JaroSimilarity(s1, s2)

    maxPrefix = 4
    If Len(s1) < maxPrefix Then maxPrefix = Len(s1)
    If Len(s2) < maxPrefix Then maxPrefix = Len(s2)

    prefixLen = 0
    For i = 1 To maxPrefix
        If LCase$(Mid$(s1, i, 1)) = LCase$(Mid$(s2, i, 1)) Then
            prefixLen = prefixLen + 1
        Else
            Exit For
        End If
    Next i

    JaroWinklerSimilarity = jaro + (CDbl(prefixLen) * scalingFactor * (1# - jaro))
End Function

Public Function JaccardSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    ' Token-based Jaccard similarity (splits on spaces)
    Dim tokens1() As String, tokens2() As String
    Dim dict As Object
    Dim i As Long
    Dim unionCount As Long, intersectCount As Long
    Dim token As String

    s1 = LCase$(Trim$(s1))
    s2 = LCase$(Trim$(s2))

    If Len(s1) = 0 And Len(s2) = 0 Then
        JaccardSimilarity = 1#
        Exit Function
    End If
    If Len(s1) = 0 Or Len(s2) = 0 Then
        JaccardSimilarity = 0#
        Exit Function
    End If

    tokens1 = Split(s1, " ")
    tokens2 = Split(s2, " ")

    Set dict = CreateObject("Scripting.Dictionary")

    ' Add all tokens from s1
    For i = LBound(tokens1) To UBound(tokens1)
        token = Trim$(tokens1(i))
        If Len(token) > 0 And Not dict.Exists(token) Then
            dict.Add token, 1  ' 1 = only in s1
        End If
    Next i

    ' Check tokens from s2
    For i = LBound(tokens2) To UBound(tokens2)
        token = Trim$(tokens2(i))
        If Len(token) > 0 Then
            If dict.Exists(token) Then
                dict(token) = 3  ' 3 = in both
            Else
                dict.Add token, 2  ' 2 = only in s2
            End If
        End If
    Next i

    unionCount = dict.Count
    intersectCount = 0
    Dim v As Variant
    For Each v In dict.Items
        If v = 3 Then intersectCount = intersectCount + 1
    Next v

    If unionCount = 0 Then
        JaccardSimilarity = 0#
    Else
        JaccardSimilarity = CDbl(intersectCount) / CDbl(unionCount)
    End If
End Function

Public Function ContainsSimilarity(ByVal s1 As String, ByVal s2 As String) As Double
    ' Returns 1 if one string contains the other, partial match based on length ratio
    Dim ls1 As String, ls2 As String
    ls1 = LCase$(Trim$(s1))
    ls2 = LCase$(Trim$(s2))

    If Len(ls1) = 0 Or Len(ls2) = 0 Then
        ContainsSimilarity = 0#
        Exit Function
    End If

    If InStr(1, ls1, ls2, vbTextCompare) > 0 Then
        ContainsSimilarity = CDbl(Len(ls2)) / CDbl(Len(ls1))
        Exit Function
    End If
    If InStr(1, ls2, ls1, vbTextCompare) > 0 Then
        ContainsSimilarity = CDbl(Len(ls1)) / CDbl(Len(ls2))
        Exit Function
    End If

    ContainsSimilarity = 0#
End Function

Public Function CombinedSimilarity(ByVal s1 As String, ByVal s2 As String, _
                                    Optional ByVal method As String = "Auto") As Double
    ' Combines multiple algorithms for best results
    Dim lev As Double, jw As Double, jac As Double, con As Double
    Dim best As Double

    Select Case LCase$(method)
        Case "levenshtein"
            CombinedSimilarity = LevenshteinSimilarity(s1, s2)
        Case "jarowinkler"
            CombinedSimilarity = JaroWinklerSimilarity(s1, s2)
        Case "jaccard"
            CombinedSimilarity = JaccardSimilarity(s1, s2)
        Case "contains"
            CombinedSimilarity = ContainsSimilarity(s1, s2)
        Case Else  ' "auto" - weighted combination
            lev = LevenshteinSimilarity(s1, s2)
            jw = JaroWinklerSimilarity(s1, s2)
            jac = JaccardSimilarity(s1, s2)
            con = ContainsSimilarity(s1, s2)

            ' Take the best score with a slight bonus for consensus
            best = lev
            If jw > best Then best = jw
            If jac > best Then best = jac
            If con > best Then best = con

            ' Weighted average biased toward best match
            CombinedSimilarity = best * 0.6 + (lev + jw + jac + con) / 4# * 0.4
    End Select
End Function

' ============================================================
' Worksheet Functions (can be used as formulas)
' ============================================================

Public Function FUZZYMATCH(ByVal value1 As String, ByVal value2 As String, _
                           Optional ByVal method As String = "Auto") As Double
    Application.Volatile False
    FUZZYMATCH = CombinedSimilarity(value1, value2, method)
End Function

Public Function FUZZYVLOOKUP(ByVal lookupValue As String, ByVal tableRange As Range, _
                              ByVal matchCol As Long, ByVal returnCol As Long, _
                              Optional ByVal threshold As Double = 0.6, _
                              Optional ByVal method As String = "Auto") As Variant
    ' Like VLOOKUP but with fuzzy matching
    ' lookupValue: the value to search for
    ' tableRange: the table to search in
    ' matchCol: column number in table to match against
    ' returnCol: column number in table to return
    ' threshold: minimum similarity (0-1)
    ' method: algorithm to use

    Application.Volatile False

    Dim bestScore As Double
    Dim bestRow As Long
    Dim score As Double
    Dim i As Long
    Dim cellValue As String

    bestScore = 0
    bestRow = 0

    For i = 1 To tableRange.Rows.Count
        cellValue = CStr(tableRange.Cells(i, matchCol).Value)
        If Len(cellValue) > 0 Then
            score = CombinedSimilarity(lookupValue, cellValue, method)
            If score > bestScore Then
                bestScore = score
                bestRow = i
            End If
        End If
    Next i

    If bestScore >= threshold And bestRow > 0 Then
        FUZZYVLOOKUP = tableRange.Cells(bestRow, returnCol).Value
    Else
        FUZZYVLOOKUP = CVErr(xlErrNA)
    End If
End Function
