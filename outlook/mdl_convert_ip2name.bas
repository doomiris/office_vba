Attribute VB_Name = "mdl_convert_ip2name"
Public Sub IP2Name()

Dim objMail As Outlook.MailItem
Set objMail = ThisOutlookSession.ActiveExplorer.Selection.Item(1)

Select Case objMail.BodyFormat
    Case olFormatPlain
        objMail.Body = ReplaceBody(objMail.Body)
    Case olFormatHTML
        objMail.HTMLBody = ReplaceBody(objMail.HTMLBody)
    Case olFormatRichText
       objMail.RTFBody = ReplaceBody(objMail.RTFBody)
End Select


End Sub

Private Function ReplaceBody(str As String)
Dim t As String
t = Replace(str, "172.24.236.60", "CCNSIF0G")

ReplaceBody = t
End Function
