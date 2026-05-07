'===============================================
' Module: String_DoubleKeywordsSplit
' DoubleKeywordsSplit
'===============================================
Public Function DoubleKeywordsSplit As Variant

Function DoubleKeywordsSplit(ByVal str As String, ByVal keyword1, ByVal keyword2) As Variant Dim resultArray() As Variant Dim currentLevel1 As Variant, currentLevel2 As Variant Dim splitParts As Variant Dim i As Long, j As Long, k As Long Dim n_UB1 As Long Dim n_UB2 As Long, n_UB2_temp As Long ReDim resultArray(0) resultArray(0) = RemoveNewLines(str) n_UB1 CountKeywordOccurrences(resultArray(0), keyword1) + 1 If n_UB1 > 1 Then ReDim currentLevel1(1 To n_UB1) splitParts = Split(resultArray(0), keyword1) For i = 1 To n_UB1 currentLevel1(i) = splitParts(i - 1) n_UB2_temp CountKeywordOccurrences(currentLevel1(i), keyword2) + 1 n_UB2 WorksheetFunction.Max(n_UB2, n_UB2_temp) Next resultArray = currentLevel1 End If If n_UB2 > 1 Then ReDim currentLevel2(1 To n_UB1, 1 To n_UB2) For i = 1 To n_UB1 n_UB2_temp CountKeywordOccurrences(resultArray(i), keyword2) + 1 splitParts = Split(resultArray(i), keyword2) For j = 1 To n_UB2_temp currentLevel2(i, j) splitParts(j - 1) Next Next resultArray = currentLevel2 End If DoubleKeywordsSplit = resultArray End Function
