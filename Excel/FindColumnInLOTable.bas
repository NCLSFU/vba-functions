'===============================================
' Module: Excel_FindColumnInLOTable
' FindColumnInLOTable
'===============================================
Public Function FindColumnInLOTable As Variant

Function FindColumnInLOTable(ByVal loTable As ListObject, ByVal keyword As String) As Long Dim headerRow As Range Dim cell As Range If loTable Is Nothing Then Exit Function End If Set headerRow = loTable.HeaderRowRange For Each cell In headerRow If cell.Value = keyword Then Exit Function End If Next cell FindColumnInLOTable = 0 End Function
