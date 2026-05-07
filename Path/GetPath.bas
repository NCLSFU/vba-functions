'===============================================
' Module: Path_GetPath
' GetPath - 获取当前工作簿路径，支持 OneDrive URL 转换
'===============================================
Public Function GetPath(ByVal basePath As String, Optional ByVal adjustLevel As Integer = 0, Optional ByVal additionalPath As String = "") As String
    ' 获取当前工作簿路径
    Dim pathParts() As String
    pathParts = Split(basePath, "\")
    
    ' 特殊路径前缀处理（例如 OneDrive）
    If Left(basePath, Len("https://d.docs.live.net/")) = "https://d.docs.live.net/" Then
        ' 将 OneDrive URL 转换为本地路径
        If Len(basePath) - Len(Replace(basePath, "/", "")) >= 4 Then
            basePath = Environ("OneDrive") & Replace(Mid(basePath, InStr(Application.Substitute(basePath, "/", "@", 4), "@")), "/", "\")
        Else
            basePath = Environ("OneDrive")
        End If
    End If
    
    ' 路径调整
    If adjustLevel > 0 Then
        Err.Raise 5, Description:="调整参数仅支持负数或 0 以减少路径层级。"
    ElseIf adjustLevel < 0 Then
        ReDim Preserve pathParts(LBound(pathParts) To UBound(pathParts) + adjustLevel)
    End If
    basePath = Join(pathParts, "\")
    
    ' 添加补充路径
    If additionalPath <> "" Then
        basePath = basePath & "\" & additionalPath
    End If
    
    ' 确保路径不以 '\' 结尾
    If Right(basePath, 1) = "\" Then
        GetPath = Left(basePath, Len(basePath) - 1)
    Else
        GetPath = basePath
    End If
End Function
