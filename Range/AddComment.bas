'===============================================
' Module: Range_AddComment
' AddComment
'===============================================
Public Sub AddComment

Sub AddComment(ByVal rng As Range, ByVal arr As Variant) If UBound(arr) < 0 Then Exit Sub End If Dim count_i As Long count_i = 0 For Each cl In rng count_i = count_i + 1 If Not cl.Comment Is Nothing Then cl.ClearComments End If cl.AddComment cl.Comment.Visible = False cl.Comment.Text Text:=arr(count_i) ' cl.Comment.Visible = True Application.CutCopyMode = False Next End Sub
