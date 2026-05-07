'===============================================
' Module: String_SplitAndProcessString
' SplitAndProcessString
'===============================================
Public Function SplitAndProcessString As Variant

Function SplitAndProcessString(inputStr As String) As Variant Dim tempStr As String Dim semicolonSplit() As String Dim commaSplit() As Variant Dim i As Long tempStr = Mid(inputStr, 5, Len(inputStr) - 6) tempStr = Replace(Replace(tempStr, vbCrLf, ""), vbLf, "") semicolonSplit = Split(tempStr, ";") ReDim commaSplit(0 To UBound(semicolonSplit)) For i = LBound(semicolonSplit) To UBound(semicolonSplit) commaSplit(i) Split(semicolonSplit(i), ",") Next i Dim finalArray() As Variant finalArray Application.WorksheetFunction.index(commaSplit, 0, 0) SplitAndProcessString = finalArray End Function
