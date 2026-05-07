'===============================================
' Module: Array_TransposeArray
' TransposeArray
'===============================================
Public Function TransposeArray As Variant

Function TransposeArray(ByVal arr As Variant) As Variant Dim numRows As Long, numCols As Long Dim transposedArr() As Variant numRows = UBound(arr, 1) - LBound(arr, 1) + 1 numCols = UBound(arr, 2) - LBound(arr, 2) + 1 ReDim transposedArr(LBound(arr, 2) To UBound(arr, 2), LBound(arr, 1) To UBound(arr, 1)) For i = LBound(arr, 1) To UBound(arr, 1) For j = LBound(arr, 2) To UBound(arr, 2) transposedArr(j, i) = arr(i, j) Next j Next i TransposeArray = transposedArr End Function
