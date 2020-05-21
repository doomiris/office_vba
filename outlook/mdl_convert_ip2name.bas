Dim map As Variant
Public Sub IP2Name()

Dim objMail As Outlook.MailItem
Set objMail = ThisOutlookSession.ActiveExplorer.Selection.Item(1)

Select Case objMail.BodyFormat
    Case olFormatPlain
        objMail.Body = ReplaceBody(objMail.Body)
    Case olFormatHTML
        objMail.HTMLBody = ReplaceBody(objMail.HTMLBody)
    Case olFormatRichText
        Dim cc As String
        cc = ReplaceBody(VBA.StrConv(objMail.RTFBody, vbUnicode))
        Dim s() As Byte
        s = VBA.StrConv(cc, vbFromUnicode)
        objMail.RTFBody = s
End Select

End Sub

Private Function ReplaceBody(ByVal str As String) As String
    str = Replace(str, "172.24.236.60", "CCNSIF0G")
    str = Replace(str, "172.24.236.61", "CCNSIA1A")
    str = Replace(str, "172.24.236.15", "CCNSIA1H")
    ReplaceBody = str
End Function

Public Sub Driver2Name()
Dim i As Long
map = Array( _
    Array("N", "\\CCNSIA1A\SolidEdge"), _
    Array("O", "\\CCNSIA1A\LDC-Projects_SHA-MZ"), _
    Array("Q", "\\CCNSIA1A\cad3d"), _
    Array("S", "\\CCNSIA1A\SEParts"), _
    Array("Y", "\\CCNSIF0G\SRDC\CCR"), _
    Array("Z", "\\CCNSIA1H\proe_stds\pro_proj_wc") _
)
If ThisOutlookSession.ActiveInspector.EditorType = olEditorWord Then
    If Not ThisOutlookSession.ActiveInspector.CurrentItem.Sent Then
        For i = LBound(map) To UBound(map)
            ThisOutlookSession.ActiveInspector.WordEditor.Content.Find.Execute FindText:=map(i)(0) & ":\", ReplaceWith:=map(i)(1) & "\", Replace:=2 'wdReplaceAll
        Next
    End If
Else
    Dim objMail As Outlook.MailItem
    Set objMail = ThisOutlookSession.ActiveInspector.CurrentItem
    
    Select Case objMail.BodyFormat
        Case olFormatPlain
            objMail.Body = ReplaceDriver(objMail.Body)
        Case olFormatHTML
            objMail.HTMLBody = ReplaceDriver(objMail.HTMLBody)
        Case olFormatRichText
           objMail.RTFBody = ReplaceDriver(objMail.RTFBody)
    End Select
End If

End Sub
Private Function ReplaceDriver(ByVal str As String) As String
    Dim i As Long
    For i = LBound(map) To UBound(map)
        str = Replace(str, map(i)(0) & ":\", map(i)(1) & "\")
    Next
    ReplaceDriver = str
End Function
