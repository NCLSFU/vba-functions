'===============================================
' Module: File_GetExcelFilePath
' GetExcelFilePath - 获取 Excel 文件路径（支持多选），需配合 GetPath 使用
'===============================================
Public Function GetExcelFilePath(ByVal basePath As String, _
                                Optional ByVal adjustLevel As Integer = 0, _
                                Optional ByVal additionalPath As String = "", _
                                Optional ByVal allowMultiSelect As Boolean = False) As Variant
    Dim targetPath As String
    Dim result As Variant
    
    targetPath = GetPath(basePath, adjustLevel, additionalPath)
    
    If Not (targetPath Like "\\*" Or targetPath Like "//*") Then
        If CurDir Like "*:\Users*\Documents" Or CurDir <> targetPath Then
            On Error Resume Next
            ChDrive Left(targetPath, 1)
            ChDir targetPath
            On Error GoTo 0
        End If
    End If
    
    result = Application.GetOpenFilename( _
        FileFilter:="Microsoft Excel文件(*.xls; *.xlsx; *.xlsm),*.xls; *.xlsx; *.xlsm", _
        MultiSelect:=allowMultiSelect)
    
    If VarType(result) = vbBoolean Then
        If Not result Then
            GetExcelFilePath = vbNullString
            Exit Function
        End If
    End If
    
    GetExcelFilePath = result
End Function
