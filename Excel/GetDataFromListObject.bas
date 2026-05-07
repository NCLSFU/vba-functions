'===============================================
' Module: Excel_GetDataFromListObject
' GetDataFromListObject
'===============================================
Public Function GetDataFromListObject As Variant

Function GetDataFromListObject(ByVal LOName As String, ByVal DataColumn As String, ByVal IndexColumn As String, ByVal IndexValue As Variant) As Variant Dim ws As Worksheet Dim LO As ListObject Dim indexCell As Range Dim dataRange As Range Dim headerRow As Long Dim rowIndex As Long Dim found As Boolean For Each ws In ThisWorkbook.Worksheets Set LO = ws.ListObjects(LOName) If Not LO Is Nothing Then found = True End If Next ws If found Then With LO Set dataRange .ListColumns(DataColumn).DataBodyRange headerRow = .HeaderRowRange.row Set indexCell .ListColumns(IndexColumn).DataBodyRange.Find(IndexValue, LookIn:=xlValues, LookAt:=xlWhole) If Not indexCell Is Nothing Then rowIndex = indexCell.row - headerRow GetDataFromListObject dataRange.Cells(rowIndex, 1).Value Else End If End With Else End If End Function
