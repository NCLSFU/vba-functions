'===============================================
' Module: String_CountKeywordOccurrences
' CountKeywordOccurrences
'===============================================
Public Function CountKeywordOccurrences As Variant

Function CountKeywordOccurrences(ByVal str As String, ByVal keyword As String) As Integer Dim count As Integer Dim pos As Integer count = 0 While pos > 0 pos = InStr(pos, str, keyword) If pos > 0 Then count = count + 1 pos = pos + Len(keyword) End If Wend CountKeywordOccurrences = count End Function
