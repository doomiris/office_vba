'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>>>>>>>> Author: Joforn          <<<<<<<<<<<<<<<<<<
'>>>>>>>> Email:  Joforn@sohu.com       <<<<<<<<<<<<<<<<<<
'>>>>>>>> QQ:   42978116        <<<<<<<<<<<<<<<<<<
'>>>>>>>> Last time : 04/11/2012        <<<<<<<<<<<<<<<<<<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Option Explicit

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>\\\\\\\\\\\\\\\\\\\\\\\ API函数定义开始  /////////////////////////<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Declare Function RegisterClipboardFormat Lib "user32.dll" Alias "RegisterClipboardFormatW" (ByVal lpString As Long) As Long
Private Declare Function OpenClipboard Lib "USER32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "USER32" () As Long
Private Declare Function SetClipboardData Lib "USER32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "USER32" () As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal flags As Long, ByVal Size As Long) As Long
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>\\\\\\\\\\\\\\\\\\\\\\\ API函数定义结束  /////////////////////////<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

Private Const CF_HDROP As Long = 15&
Private Const DROPEFFECT_COPY As Long = 1
Private Const DROPEFFECT_MOVE As Long = 2
Private Const GMEM_ZEROINIT As Long = &H40
Private Const GMEM_MOVEABLE As Long = &H2
Private Const GMEM_DDESHARE As Long = &H2000


'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>\\\\\\\\\\\\\\\\\\\\\\\  结构定义开始  /////////////////////////<
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type dropFiles
  pFiles As Long
  pt  As POINTAPI
  fNC As Long
  fWide As Long
End Type
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'>\\\\\\\\\\\\\\\\\\\\\\\  结构定义结束   ////////////////////////<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

Public Function CutOrCopyFiles(FileList As Variant, Optional ByVal CopyMode As Boolean = True) As Boolean
Dim uDropEffect As Long, i As Long
Dim dropFiles As dropFiles
Dim uGblLen As Long, uDropFilesLen As Long
Dim hGblFiles As Long, hGblEffect As Long
Dim mPtr  As Long
Dim FileNames As String

If OpenClipboard(0) Then
  EmptyClipboard
  FileNames = GetFileListString(FileList)
  If Len(FileNames) Then
  uDropEffect = RegisterClipboardFormat(StrPtr("Preferred DropEffect"))
  hGblEffect = GlobalAlloc(GMEM_ZEROINIT Or GMEM_MOVEABLE Or GMEM_DDESHARE, Len(uDropEffect))
  mPtr = GlobalLock(hGblEffect)
  i = IIf(CopyMode, DROPEFFECT_COPY, DROPEFFECT_MOVE)
  CopyMemory ByVal mPtr, i, Len(i)
  GlobalUnlock hGblEffect
  
  uDropFilesLen = Len(dropFiles)
  With dropFiles
  .pFiles = uDropFilesLen
  .fWide = CLng(True)
  End With
  uGblLen = uDropFilesLen + Len(FileNames) * 2 + 8
  hGblFiles = GlobalAlloc(GMEM_ZEROINIT Or GMEM_MOVEABLE Or GMEM_DDESHARE, uGblLen)
  mPtr = GlobalLock(hGblFiles)
  CopyMemory ByVal mPtr, dropFiles, uDropFilesLen
  CopyMemory ByVal (mPtr + uDropFilesLen), ByVal StrPtr(FileNames), LenB(FileNames)
  GlobalUnlock hGblFiles
  SetClipboardData CF_HDROP, hGblFiles
  End If
  CloseClipboard
End If
End Function

Private Function GetFileListString(FileList As Variant) As String
Dim i As Long

On Error GoTo GetFileListStringLOOP
Select Case VarType(FileList)
  Case vbString
  GetFileListString = Trim$(FileList)
  Case &H2008
  For i = LBound(FileList) To UBound(FileList)
  FileList(i) = Trim$(FileList(i))
  If Len(FileList(i)) Then GetFileListString = GetFileListString & FileList(i) & vbNullChar
  Next i
  If Len(GetFileListString) Then GetFileListString = Left$(GetFileListString, Len(GetFileListString) - 1)
End Select
GetFileListStringLOOP:

End Function


