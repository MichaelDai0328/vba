Attribute VB_Name = "Module1"
Sub BatchConvertWorkBookToCSV()
Application.DisplayAlerts = False
Application.ScreenUpdating = False

    Dim fDialog As FileDialog
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Dim vrtSelectedItem As Variant
    Dim wkBook As Workbook
    Dim showFolder  As Boolean
    showFolder = False
    With fDialog
        .Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm", 1
        If .Show = -1 Then
            For Each vrtSelectedItem In .SelectedItems
                '如果选择了本工作簿则跳过
                If InStrRev(vrtSelectedItem, ThisWorkbook.Name) = 0 Then
                    On Error Resume Next
                    Set wkBook = Application.Workbooks.Open(vrtSelectedItem, ReadOnly:=True, Password:="")
                    '跳过设置打开密码的工作簿
                    If Not wkBook Is Nothing Then
                       '跳过隐藏的工作簿
                       If Windows(wkBook.Name).Visible = True Then
                       showFolder = True
                       '转换开始
                       wkBook.SaveAs FileFormat:=xlWorkbookNormal, Filename:= _
                          Left(vrtSelectedItem, InStrRev(vrtSelectedItem, ".") - 1) & ".xls" _
                          , CreateBackup:=False
                       wkBook.Close , savechanges = False
                       Else
                       wkBook.Close , savechanges = False
                       End If
                    End If
               End If
            Next vrtSelectedItem
            If showFolder Then Call Shell("explorer.exe " & Left(fDialog.SelectedItems(1), _
                InStrRev(fDialog.SelectedItems(1), "\")), vbMaximizedFocus)
        End If
    End With
    
    Set fDialog = Nothing
Application.ScreenUpdating = True
Application.DisplayAlerts = True
End Sub



