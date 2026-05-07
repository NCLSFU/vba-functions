'===============================================
' Module: Path_CreatePath
' CreatePath - 递归创建目录路径，确保每一级目录都存在
'===============================================
Public Sub CreatePath(ByVal fullPath As String)
    Dim pathParts() As String
    Dim currentPath As String
    
    pathParts = Split(Application.Trim(fullPath), "\")
    
    For i = 0 To UBound(pathParts)
        currentPath = currentPath & pathParts(i) & "\"
        If Not Dir(currentPath, vbDirectory) <> vbNullString Then
            MkDir currentPath
        End If
    Next i
End Sub
