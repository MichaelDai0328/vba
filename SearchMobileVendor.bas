Attribute VB_Name = "Module1"
Private Sub search_vendor()
    Dim temp
    On Error Resume Next
    With CreateObject("Microsoft.XMLHTTP")
        For p = 2 To Range("d65536").End(xlUp).Row
            .Open "GET", "http://www.hiphop8.com/nub/" & Cells(p, 4) & ".html", True
            .Send
            Do Until .ReadyState = 4
                DoEvents
            Loop
            temp = Split(StrConv(.responseBody, vbUnicode, &H804), "class=T>")
            vendor = Split(Split(temp(3), "<")(0), " ")(0)
            Cells(p, "F") = Split(Split(temp(1), "<")(0), " ")(1)
            Cells(p, "G") = Mid(vendor, 1, 2)
        Next p
    End With
End Sub

