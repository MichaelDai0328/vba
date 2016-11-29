Attribute VB_Name = "Module1"
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

    Dim lngres As Long
    
    On Error Resume Next

    'check attachment

    If InStr(1, Item.Body, "attach") <> 0 Then

    If Item.Attachments.Count = 0 Then

    Application.Explorers(1).Activate

    lngres = MsgBox("no attachment found !" & Chr(10) & "send anyway?", vbYesNo + vbDefaultButton2 + vbQuestion, "tip")

    If lngres = vbNo Then

    Cancel = True

    Item.Display

    Exit Sub

    End If
    
    End If

    End If

   

    'check subject

    If Item.Subject = "" Then

    Application.Explorers(1).Activate

    lngres = MsgBox("No subject!" & Chr(10) & "send anyway?", vbYesNo + vbDefaultButton2 + vbQuestion, "tip")

    If lngres = vbNo Then

    Cancel = True

    Item.Display

    Exit Sub

    End If

    End If


    'check sender

    If InStr(1, Item.SentOnBehalfOfName, "SDC ES&S") <> 1 Then

    Application.Explorers(1).Activate

    lngres = MsgBox("Didn't send on behalf of FCYESSM" & Chr(10) & "send anyway?", vbYesNo + vbDefaultButton2 + vbQuestion, "tip")

    If lngres = vbNo Then

    Cancel = True

    Item.Display

    Exit Sub

    End If
  
    End If

  
End Sub

