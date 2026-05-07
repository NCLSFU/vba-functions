'===============================================
' Module: Path_CreatFilePath
' CreatFilePath - 创建带年月/公司简称的文件路径（含目录创建）
'===============================================
Function CreatFilePath(ByVal 主路径 As String, ByVal 表类 As String, ByVal 表名 As String, ByVal 公司简称 As String, ByVal 年 As Integer, ByVal 月 As Integer, Optional ByVal 文件格式 As String = ".xlsx") As String
    Dim arr_FP1(), arr_FP2()
    ReDim arr_FP1(1 To 5)
    ReDim arr_FP2(1 To 4)
    
    arr_FP1(1) = 主路径
    arr_FP1(2) = 表类
    arr_FP1(3) = 表名
    arr_FP1(4) = 年 & "-" & 月
    arr_FP1(5) = 表类 & "_" & 表名 & "_" & 公司简称 & "_" & 年 & "-" & 月 & 文件格式
    
    For i = 1 To 4
        arr_FP2(i) = arr_FP1(i)
    Next
    
    Dim tempFilePath As String
    tempFilePath = Join(arr_FP2, "\")
    
    Call CreatePath(tempFilePath)
    CreatFilePath = Join(arr_FP1, "\")
End Function
