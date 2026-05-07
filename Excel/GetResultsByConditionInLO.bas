'===============================================
' Module: Excel_GetResultsByConditionInLO
' GetResultsByConditionInLO
'===============================================
Public Function GetResultsByConditionInLO As Variant

Function GetResultsByConditionInLO(ByVal LO As ListObject, ByVal ResultCol As String, ByVal ConditionCol As String, ByVal SearchText As String) As Variant Dim Results() Dim i As Long Dim MatchedRowsCount As Long For i = 1 To LO.ListRows.count If LO.ListColumns(ConditionCol).DataBodyRange(i) Like SearchText Then Results = ArrayAppend(Results, LO.ListColumns(ResultCol).DataBodyRange(i)) End If Next i GetResultsByConditionInLO = Results End Function
