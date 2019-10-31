Attribute VB_Name = "mdl_openSelectedPDF"
Public Sub openSelectedPdf()

Dim objMail As Outlook.MailItem
Set objMail = ThisOutlookSession.ActiveExplorer.Selection.Item(1)
Dim wordSelected As String
wordSelected = objMail.GetInspector.WordEditor.Application.Selection.Text
wordSelected = Trim(wordSelected)
wordSelected = Replace(Replace(wordSelected, Chr(10), ""), Chr(13), "")             'ȥ�����з�

Dim myPDFstore As String

Dim p As Variant
p = Split(GetSetting("Domisoft", "Config", "PDF_Store", ""), "|")

Dim filename As String
Dim done As Boolean
Dim fullName As String
done = False
For i = LBound(p) To UBound(p)

    myPDFstore = p(i)
    
    filename = wordSelected
    
    If InStr(1, filename, Chr(10), vbTextCompare) > 0 Then
        filename = Split(filename, Chr(10))(0)           ' TODO һ���ﺬ�ж���ļ���
    End If
    If Len(filename) = 8 And Left(filename, 1) = 8 Then filename = "00" & filename    '���00����
    
    fullName = myPDFstore & "\" & filename & ".pdf"
    
    If IsFileExists(fullName) Then
        Shell "explorer.exe " & fullName
        done = True
        Exit For
    End If
Next

If Not done Then
    fullName = SearchFor(wordSelected)
    If Len(Trim(fullName)) > 0 Then
        Shell "explorer.exe " & fullName
        done = True
    End If
End If

If Not done Then MsgBox "file not found: " & fullName, vbOKOnly, "File Not Found"
End Sub
Public Function IsFileExists(ByVal strFileName As String) As Boolean
    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function
Private Function SearchFor(filenameKey As String) As String
    myPath$ = "S:\Cabinet\QHCͼֽ��\"
    myfile$ = "*" & filenameKey & "*"

    Set wshell = CreateObject("Wscript.Shell") 'VBA����Dos����
    ar = Split(wshell.exec("cmd /c dir /a-d /b /s " & Chr(34) & myPath & myfile & Chr(34)).StdOut.ReadAll, vbCrLf) '�����ĵ������ļ���
    
'    For i = 0 To UBound(ar)
'        If Len(Trim(ar(i))) > 0 Then Debug.Print i, ar(i)
'    Next
    
    SearchFor = ar(0) 'ѡȡ��һ��

End Function

