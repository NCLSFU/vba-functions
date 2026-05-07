'===============================================
' Module: Excel_ClearListObjectDataRange
' ClearListObjectDataRange
'===============================================
Public Function ClearListObjectDataRange As Variant

Function ClearListObjectDataRange(ByVal LO As ListObject) If Not LO Is Nothing Then If LO.ListRows.count > 0 Then Dim i As Long For i = LO.ListRows.count To 1 Step -1 LO.ListRows(i).Delete Next i End If Else End If End Function
