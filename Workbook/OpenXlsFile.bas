'===============================================
' Module: Workbook_OpenXlsFile
' OpenXlsFile
'===============================================
Public Function OpenXlsFile As Variant

Function OpenXlsFile(ByVal FilePath, Optional ByVal OpenDisplay As Boolean = True) As Workbook On Error Resume Next With Workbooks(Mid(FilePath, InStrRev(FilePath, "\") + 1)) End With On Error GoTo 0 On Error Resume Next If OpenDisplay Then Set OpenXlsFile Workbooks.Open(FilePath) Else Set OpenXlsFile GetObject(FilePath) End If On Error GoTo 0 If OpenXlsFile Is Nothing Then End If End Function
