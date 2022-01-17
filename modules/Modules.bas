Attribute VB_Name = "Modules"
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpstring As Any, ByVal lpFileName As String) As Long
Public OpFile As String

    Public Declare Function ShellExecute Lib _
     "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal _
     lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long

Public Type BrowseInfo
hWndOwner As Long
pIDLRoot As Long
pszDisplayName As Long
lpszTitle As Long
ulFlags As Long
lpfnCallback As Long
lParam As Long
iImage As Long
End Type

Public Const BIF_RETURNONLYFSDIRS = 1
Public Const MAX_PATH = 260

Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" _
(ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" _
(lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" _
(ByVal pidList As Long, ByVal lpBuffer As String) As Long

Sub Main()
frmSplash.Show
OpFile = App.Path & "\option.ini"
Load frmMain
frmMain.Show
Unload frmSplash
End Sub

Public Function Loadini(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, Optional KeyDefault As String)
    Dim rc As Integer
    Dim KeyValue As String
    Dim sectionname As String
    Dim shuzhi1 As String
    Dim shuzhi2 As String
    Dim TR
    Dim z As Integer
    KeyValue = String$(255, 0)
    shuzhi1 = String$(255, 0)
    rc = GetPrivateProfileString(AppName, KeyName, KeyDefault, shuzhi1, Len(shuzhi1), FileName)
   TR = Left$(shuzhi1, rc)
   For i = 1 To Len(TR)
   X = Right(Left(TR, i), 1)
   Y = Asc(X)
   If Y = 0 Then
   z = z + 1
   End If
   Next
   TR = Left(TR, Len(TR) - z)
Loadini = TR
End Function

Public Sub Saveini(ByVal FileName As String, ByVal AppName As String, ByVal KeyName As String, ByVal KeyValue As String)
WritePrivateProfileString AppName, KeyName, KeyValue, FileName
End Sub

Public Function BrowseForFolder(hWndOwner As Long, sPrompt As String) As String

Dim iNull As Integer
Dim lpIDList As Long
Dim lResult As Long
Dim sPath As String
Dim udtBI As BrowseInfo


With udtBI
.hWndOwner = hWndOwner
.lpszTitle = lstrcat(sPrompt, "")
.ulFlags = BIF_RETURNONLYFSDIRS
End With


lpIDList = SHBrowseForFolder(udtBI)
If lpIDList Then
sPath = String$(MAX_PATH, 0)
lResult = SHGetPathFromIDList(lpIDList, sPath)
Call CoTaskMemFree(lpIDList)
iNull = InStr(sPath, vbNullChar)
If iNull Then
sPath = Left$(sPath, iNull - 1)
End If
End If

BrowseForFolder = sPath
End Function


