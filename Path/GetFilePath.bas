'===============================================
' Module: Path_GetFilePath
' GetFilePath - 根据表类、表名、公司、年、月获取指定文件的路径
'===============================================
Function GetFilePath(ByVal 主路径 As String, ByVal 表类 As String, ByVal 表名 As String, ByVal 公司简称 As String, ByVal 年 As Integer, ByVal 月 As Integer) As String
    Dim arr_FP1(), arr_FP2()
    ReDim arr_FP1(1 To 5)
    ReDim arr_FP2(1 To 4)
    
    arr_FP1(1) = 主路径
    arr_FP1(2) = 表类
    arr_FP1(3) = 表名
    arr_FP1(4) = 年 & "-" & 月
    arr_FP1(5) = 表类 & "_" & 表名 & "_" & 公司简称 & "_" & 年 & "-" & 月 & ".*"
    
    For i = 1 To 4
        arr_FP2(i) = arr_FP1(i)
    Next
    
    Dim tempFilePath As String
    tempFilePath = Join(arr_FP1, "\")
    
    If Len(Dir(tempFilePath)) > 0 Then
        GetFilePath = Join(arr_FP2, "\") & "\" & Dir(tempFilePath)
    Else
        GetFilePath = ""
    End If
End Function
