'===============================================
' Module: Excel_AppendListData
' AppendListData
'===============================================
Public Function AppendListData As Variant

Function AppendListData(ByVal dstListObj As ListObject, ByVal srcListObj As ListObject) As Boolean Dim srcRange As Range Dim dstRange As Range Dim arrData() As Variant Dim i As Long, j As Long Set srcRange = srcListObj.DataBodyRange ReDim arrData(1 To srcRange.Rows.count, 1 To srcRange.columns.count) For i = 1 To srcRange.Rows.count For j = 1 To srcRange.columns.count arrData(i, j) = srcRange.Cells(i, j).Value Next j Next i Set dstRange = dstListObj.DataBodyRange If Not dstRange Is Nothing Then If dstRange.Rows.count > 0 Then Set dstRange dstRange.Cells(dstRange.Rows.count, 1).Offset(1, 0) Else Set dstRange dstListObj.ListColumns(1).Range.Cells(1, 1) End If Else AppendListData = False Exit Function End If dstRange.Resize(UBound(arrData, 1), UBound(arrData, 2)).Value = arrData AppendListData = True End Function
