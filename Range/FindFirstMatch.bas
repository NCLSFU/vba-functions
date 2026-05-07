'===============================================
' Module: Range_FindFirstMatch
' FindFirstMatch
'===============================================
Public Function FindFirstMatch As Variant

Function FindFirstMatch(ByVal keyword As String, ByVal searchRange As Range, Optional ByVal exactMatch As Boolean = False) As Range Dim cell As Range Dim foundCell As Range Dim lookAtType As XlLookAt If exactMatch Then lookAtType = xlWhole Else lookAtType = xlPart End If Set foundCell searchRange.Find(What:=keyword, _ LookIn:=xlValues, _ LookAt:=lookAtType, _ SearchOrder:=xlByRows, _ SearchDirection:=xlNext, _ MatchCase:=False) If Not foundCell Is Nothing Then Set FindFirstMatch = foundCell Else Set FindFirstMatch = Nothing End If End Function
