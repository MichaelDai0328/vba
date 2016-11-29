Attribute VB_Name = "Module1"
Sub style()
    Dim rowNum As Long
    Dim colNum As Long
    Dim i As Integer
    
    colNum = Sheet1.Range("a2").End(xlToRight).Column
    rowNum = Sheet1.Range("A65536").End(xlUp).Row
    
    With Sheet1.Range(Cells(2, 1), Cells(rowNum, colNum)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With

    With Sheet1.Range(Cells(2, 1), Cells(rowNum, colNum)).Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    For i = 2 To rowNum
    
    If i Mod 2 = 0 Then
        Sheet1.Range(Cells(i, 1), Cells(i, colNum)).Interior.ColorIndex = 15
    Else
        Sheet1.Range(Cells(i, 1), Cells(i, colNum)).Interior.ColorIndex = 0
    End If
    Next
    
End Sub


    
Private Sub Worksheet_Change(ByVal Target As Range)

Call style

End Sub


