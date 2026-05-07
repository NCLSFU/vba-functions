'===============================================
' Module: Excel_ClearListObjecDataBodyRange
' ClearListObjecDataBodyRange
'===============================================
Public Function ClearListObjecDataBodyRange As Variant

Function ClearListObjecDataBodyRange(ByVal LO As ListObject, Optional ByVal deletAll As Boolean = True) ' V2 If Not LO Is Nothing Then If deletAll Then If Not LO.DataBodyRange Is Nothing Then LO.DataBodyRange.Delete End If Else If Not LO.DataBodyRange Is Nothing Then LO.DataBodyRange.ClearContents End If End If Else End If End Function
