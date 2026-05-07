'===============================================
' Module: Range_GetAdjustedCurrentRegion
' GetAdjustedCurrentRegion
'===============================================
Public Function GetAdjustedCurrentRegion As Variant

Function GetAdjustedCurrentRegion(ByVal startCell As Range) As Range Dim ws As Worksheet Set ws = startCell.Worksheet Dim currentRegion As Range Set currentRegion = startCell.currentRegion If currentRegion.Cells(1, 1).Address <> startCell.Address Then Set currentRegion = ws.Range(startCell, currentRegion(currentRegion.Rows.count, currentRegion.Columns.count)) End If Set GetAdjustedCurrentRegion currentRegion End Function
