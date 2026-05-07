'===============================================
' Module: Array_ArrayAppend
' ArrayAppend
'===============================================
Public Function ArrayAppend As Variant

Function ArrayAppend(ByVal arr As Variant, ByVal element As Variant, Optional ByVal lb As Long = 1) As Variant Dim isArr As Boolean isArr = IsArray(arr) Dim resultArr() As Variant Dim i As Long, arrLen As Long, elemArrLen As Long If isArr Then If IsArrayEmpty(arr) Then isArr = False Else arrLen = UBound(arr) - LBound(arr) + 1 End If Else Stop Exit Function End If If IsArray(element) Then If Not IsArrayEmpty(element) Then elemArrLen = UBound(element) - LBound(element) + 1 End If ElseIf Not IsEmpty(elemment) Then elemArrLen = 1 End If arrLen = arrLen + elemArrLen ReDim resultArr(lb To lb + arrLen - 1) If isArr Then For i = LBound(arr) To UBound(arr) resultArr(i) = arr(i) Next i End If If IsArray(element) Then For i = LBound(element) To UBound(element) resultArr(lb + arrLen - elemArrLen + i) = element(i) Next i ElseIf Not IsEmpty(elemment) Then resultArr(lb + arrLen - 1) = element End If ArrayAppend = resultArr End Function
