'===============================================
' Module: Excel_AdjustAndPasteArrayToLO
' AdjustAndPasteArrayToLO
'===============================================
Public Sub AdjustAndPasteArrayToLO

Sub AdjustAndPasteArrayToLO(ByVal LO As ListObject, ByVal targetCell As Range, ByVal arr As Variant) If UBound(arr) < 0 Then Exit Sub End If Dim targetRow As Long, targetColumn As Long targetRow = targetCell.row targetColumn = targetCell.Column Dim numRows As Long, numCols As Long numRows = UBound(arr, 1) - LBound(arr, 1) + 1 numCols = UBound(arr, 2) - LBound(arr, 2) + 1 Dim currentRange As Range Set currentRange = LO.Range Dim lastRow As Long, lastColumn As Long lastRow = targetRow - 1 lastColumn = targetColumn - 1 For i = 0 To numRows - 1 For j = 0 To numCols - 1 If i + targetRow > LO.Range.Rows.count + LO.HeaderRowRange.row - 1 Or _ j + targetColumn > LO.Range.columns.count + LO.HeaderRowRange.Column - 1 Then LO.Parent.Activate LO.Resize Range(LO.Range.Cells(1, 1), _ Cells(WorksheetFunction.Max(i + targetRow, LO.Range.Rows.count + LO.HeaderRowRange.row - 1), _ WorksheetFunction.Max(j + targetColumn, LO.Range(1, LO.Range.columns.count).Column))) End If targetCell.Offset(i, j).Value arr(LBound(arr, 1) + i, LBound(arr, 2) + j) ' If i + targetRow > lastRow Then lastRow = i + targetRow ' If j + targetColumn > lastColumn Then lastColumn = j + targetColumn Next j Next i End Sub
