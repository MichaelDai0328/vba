Attribute VB_Name = "Module1"
Sub ChangeShtName()
    Dim rng As Range
    'On Error Resume Next
    Set rng = Application.InputBox(prompt:="��ѡȡ�����뵥Ԫ������", Type:=8)
    rng.Select
    i = 4
    For Each cell In rng
        Sheets(i).Name = cell.Value2
        i = i + 1
    Next
End Sub
