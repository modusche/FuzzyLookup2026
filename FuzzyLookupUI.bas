Attribute VB_Name = "FuzzyLookupUI"
Option Explicit

' ============================================================
' Fuzzy Lookup UI - Ribbon callback and form launcher
' ============================================================

Public Sub ShowFuzzyLookupForm()
    Dim frm As FuzzyLookupForm
    Set frm = New FuzzyLookupForm
    frm.Show
End Sub

Public Sub ShowAbout()
    MsgBox "Fuzzy Lookup Add-In for Excel" & vbCrLf & _
           "Version 2.0.0 (64-bit compatible)" & vbCrLf & vbCrLf & _
           "Performs fuzzy matching of textual data between two Excel tables." & vbCrLf & vbCrLf & _
           "Algorithms:" & vbCrLf & _
           "  - Levenshtein Distance" & vbCrLf & _
           "  - Jaro-Winkler Similarity" & vbCrLf & _
           "  - Jaccard Token Similarity" & vbCrLf & _
           "  - Contains Matching" & vbCrLf & _
           "  - Auto (weighted combination)" & vbCrLf & vbCrLf & _
           "Worksheet Functions:" & vbCrLf & _
           "  =FUZZYMATCH(A1, B1)" & vbCrLf & _
           "  =FUZZYVLOOKUP(A1, Table, 1, 2, 0.6)", _
           vbInformation, "About Fuzzy Lookup"
End Sub

Public Function GetNamedRangesAndTables() As Collection
    ' Returns a collection of all named ranges and tables in the active workbook
    Dim col As New Collection
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim nm As Name

    On Error Resume Next

    ' Add tables
    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            col.Add lo.Name & " (Table on " & ws.Name & ")", lo.Name
        Next lo
    Next ws

    ' Add named ranges
    For Each nm In ActiveWorkbook.Names
        If Not nm.Name Like "_*" Then
            col.Add nm.Name & " (Named Range)", "NR_" & nm.Name
        End If
    Next nm

    On Error GoTo 0
    Set GetNamedRangesAndTables = col
End Function

Public Function RangeFromTableOrName(ByVal tableName As String) As Range
    ' Resolves a table name or named range to a Range object
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next

    ' Try as table first
    For Each ws In ActiveWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName Then
                Set RangeFromTableOrName = lo.Range
                Exit Function
            End If
        Next lo
    Next ws

    ' Try as named range
    Set RangeFromTableOrName = ActiveWorkbook.Names(tableName).RefersToRange

    On Error GoTo 0
End Function
