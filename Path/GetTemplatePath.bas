'===============================================
' Module: Path_GetTemplatePath
' GetTemplatePath - 根据输入的模板类、模板名、子类获取指定文件的路径
'===============================================
Function GetTemplatePath(ByVal 主路径 As String, ByVal 模板类 As String, ByVal 模板名 As String, ByVal 子类 As String) As String
    Dim arr_FP1(), arr_FP2()
    ReDim arr_FP1(1 To 4)
    ReDim arr_FP2(1 To 3)
    
    arr_FP1(1) = 主路径
    arr_FP1(2) = 模板类
    arr_FP1(3) = 模板名
    arr_FP1(4) = 模板类 & "_" & 模板名 & "_" & 子类 & ".*"
    
    For i = 1 To 3
        arr_FP2(i) = arr_FP1(i)
    Next
    
    Dim tempFilePath As String
    tempFilePath = Join(arr_FP1, "\")
    
    If Len(Dir(tempFilePath)) > 0 Then
        GetTemplatePath = Join(arr_FP2, "\") & "\" & Dir(tempFilePath)
    Else
        GetTemplatePath = ""
    End If
End Function
