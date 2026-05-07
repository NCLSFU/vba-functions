'===============================================
' Module: Array_IsArrayEmpty
' IsArrayEmpty
'===============================================
Public Function IsArrayEmpty As Variant

Function IsArrayEmpty(arr) As Boolean If Not IsArray(arr) Then Else On Error GoTo ErrorHandler If UBound(arr) >= LBound(arr) Then IsArrayEmpty = False Exit Function End If ErrorHandler: IsArrayEmpty = True End Function
