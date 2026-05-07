'===============================================
' Module: Path_FindExcelFilePathWithKeyword
' FindExcelFilePathWithKeyword - 查找目录下包含关键字的 Excel 文件路径
'===============================================
Function FindExcelFilePathWithKeyword(ByVal str As String) As Variant
    Dim filePath As String
    Dim fileName As String
    Dim fullFileName As String
    Dim i As Integer
    
    FindExcelFilePathWithKeyword = ""
    
    filePath = GetPath(ThisWorkbook.Path)
    
    fileName = Dir(filePath & "\*.xl*")
    Do While fileName <> ""
        If fileName Like str Then
            FindExcelFilePathWithKeyword = filePath & "\" & fileName
            Exit Do
        End If
        fileName = Dir
    Loop
    
    If FindExcelFilePathWithKeyword = "" Then MsgBox "未在当前路径找到符合规则（" & str & "）的文件"
End Function
