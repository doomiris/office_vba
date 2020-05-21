#If VBA7 Then
    Declare PtrSafe Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#Else
    Declare Function GetKeyState Lib "USER32" (ByVal vKey As Long) As Integer
#End If


Public Sub openSelectedPdf()
Dim copyFile As Boolean
copyFile = GetKeyState(vbKeyControl) < 0       '复制到剪切板

Dim objMail As Outlook.MailItem
Set objMail = ThisOutlookSession.ActiveExplorer.Selection.Item(1)
Dim wordSelected As String
wordSelected = objMail.GetInspector.WordEditor.Application.Selection.Text
wordSelected = Trim(wordSelected)
wordSelected = Replace(Replace(wordSelected, Chr(10), ""), Chr(13), "")             '去除换行符

Dim myPDFstore As String

Dim p As Variant
p = Split(GetSetting("Domisoft", "Config", "PDF_Store", ""), "|")

Dim filename As String
Dim done As Boolean
Dim fullName As String
done = False

'Excel.Application.cursor = xlWait '修改鼠标为等待
On Error GoTo eHand


For i = LBound(p) To UBound(p)

    myPDFstore = p(i)
    
    filename = wordSelected
    
    If Len(filename) = 0 Then Exit Sub
    
    If InStr(1, filename, Chr(10), vbTextCompare) > 0 Then
        filename = Split(filename, Chr(10))(0)           ' TODO 一格里含有多个文件名
    End If
    If Len(filename) = 8 And Left(filename, 1) = 8 Then filename = "00" & filename    '解决00问题
    If Len(filename) = 11 And Left(filename, 1) = "H" Then filename = Right(filename, Len(filename) - 1) '解决H问题
    
    If Right(myPDFstore, 1) = "\" Then myPDFstore = Left(myPDFstore, Len(myPDFstore) - 1) '解决path以\结尾问题
    
    'filename = myPDFstore & "\" & filename & ".pdf"
    
    filename = SearchFor(filename, myPDFstore)
    
    If IsFileExists(filename) And Len(filename) > 0 Then
    
        If copyFile Then   '如果按下CTRL键
            send2Clipboard filename         '复制到剪切板
        Else                                    '否则
            Shell "explorer.exe " & filename    '直接打开
        End If

        done = True
        
        Exit For
    End If
Next
If Not done Then MsgBox "file not found: " & filename, vbOKOnly, "File Not Found"
eHand:
'Excel.Application.cursor = xlDefault '恢复鼠标
End Sub
Public Function IsFileExists(ByVal strFileName As String) As Boolean
    If Dir(strFileName, 16) <> Empty Then
        IsFileExists = True
    Else
        IsFileExists = False
    End If
End Function
Private Function SearchFor(filenameKey As String, Optional libPath As String) As String
    If Len(libPath) = 0 Then libPath = "S:\Cabinet\QHC图纸库\"
    
    If Right(libPath, 1) <> "\" Then libPath = libPath & "\"
    
    myfile$ = "*" & filenameKey & "*"
    toFolder$ = GetSetting("Domisoft", "Config", "SE_Output", "d:\")
    tempOut$ = toFolder & "\out.txt"
    
    'Dim wshell As New WshShell         '直接CreateObject不用引用
    Set wshell = CreateObject("Wscript.Shell") 'VBA调用Dos命令
    
    searchSTR$ = "cmd /c dir /a-d /b /s " & Chr(34) & libPath & myfile & Chr(34) & " > " & tempOut
    
    ' Pass 0 as the second parameter to hide the window...
    wshell.Run searchSTR, 0, True
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next    '读空文件时会出错.
    ss = fso.OpenTextFile(tempOut).ReadAll()
    fso.DeleteFile tempOut
    On Error GoTo 0
    ar = Split(ss, vbCrLf) '所有文档含子文件夹
    
'    For i = 0 To UBound(ar)
'        If Len(Trim(ar(i))) > 0 Then Debug.Print i, ar(i)
'    Next
    
    If UBound(ar) > -1 Then
        SearchFor = ar(0) '选取第一条
    Else
        SearchFor = ""
    End If
End Function

Public Sub send2Clipboard(filename As String)
    Dim succ As Boolean
    succ = CutOrCopyFiles(filename, True)
    Debug.Print filename
End Sub

