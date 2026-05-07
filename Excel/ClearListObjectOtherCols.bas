'===============================================
' Module: Excel_ClearListObjectOtherCols
' ClearListObjectOtherCols
'===============================================
Public Function ClearListObjectOtherCols As Variant

Function ClearListObjectOtherCols(ByVal LO As ListObject, ByVal n_reserved As Long) If Not LO Is Nothing Then If LO.ListColumns.count > n_reserved Then Dim i As Long For i = LO.ListColumns.count To n_reserved + 1 Step -1 LO.ListColumns(i).Delete Next i End If Else End If End Function
