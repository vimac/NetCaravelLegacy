VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "选项"
   ClientHeight    =   4425
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   5505
   Icon            =   "frmOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame 
      Height          =   3255
      Index           =   1
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton cmdAppsEdit 
         Caption         =   "运行"
         Height          =   375
         Left            =   1560
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin MSComctlLib.ListView lstApps 
         Height          =   2535
         Left            =   0
         TabIndex        =   15
         Top             =   720
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         GridLines       =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "程序名"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "命令"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.CommandButton cmdAppsDel 
         Caption         =   "删除"
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton cmdAppsAdd 
         Caption         =   "添加"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame 
      Height          =   3255
      Index           =   0
      Left            =   240
      TabIndex        =   3
      Top             =   420
      Width           =   4935
      Begin VB.CommandButton cmdFolder 
         Caption         =   "浏览"
         Height          =   255
         Left            =   4080
         TabIndex        =   12
         Top             =   2280
         Width           =   735
      End
      Begin VB.TextBox txtFav 
         Height          =   270
         Left            =   120
         TabIndex        =   10
         Text            =   "c:\"
         Top             =   2280
         Width           =   3855
      End
      Begin VB.CheckBox chkShowSideBar 
         Caption         =   "显示边条"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtHome 
         Height          =   270
         Left            =   120
         TabIndex        =   7
         Text            =   "http://"
         Top             =   1560
         Width           =   4695
      End
      Begin VB.CheckBox chkAllowNewWindow 
         Caption         =   "允许网页打开新标签框"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3255
      End
      Begin VB.CheckBox chkXPStyle 
         Caption         =   "XP样式的菜单(9x系统上面可能会出错)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "书签路径"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "首页"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   4200
      TabIndex        =   2
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tabOption 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6588
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "常规"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "外接程序"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAppsAdd_Click()
Dim i As Integer, x As String
On Error GoTo Errors
i = lstApps.ListItems.Count + 1

    frmMain.CDLG.DialogTitle = "选择程序"
    frmMain.CDLG.Filter = "程序文件|*.exe|"
    frmMain.CDLG.ShowOpen
    x = InputBox("请输入程序名", "Plz input...", "应用程序")
    Saveini OpFile, "apps", "app" & Str(i), x & "," & frmMain.CDLG.FileName
    LoadApps
    
Exit Sub

Errors:

End Sub

Private Sub cmdAppsDel_Click()
Saveini OpFile, "apps", "app" & Str(lstApps.SelectedItem.Index), ""
LoadApps
End Sub

Private Sub cmdAppsEdit_Click()
Shell lstApps.SelectedItem.ListSubItems(1).Text
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
Saveini OpFile, "navigator", "home", txtHome.Text
Saveini OpFile, "navigator", "favorite", txtFav.Text
Saveini OpFile, "navigator", "enablenewwindow", Me.chkAllowNewWindow.Value
Saveini OpFile, "window", "xp", Me.chkXPStyle.Value
Saveini OpFile, "window", "sidebar", Me.chkShowSideBar.Value

MsgBox "一些设置需要重新启动nc3才能够生效"
Unload Me
End Sub

Sub LoadApps()
On Error Resume Next
Dim i As Integer, x As Integer, path As String, aname As String
lstApps.ListItems.Clear
For i = 1 To 40
    aname = Loadini(OpFile, "apps", "app" & Str(i))
    If aname <> "" Then
        x = InStr(1, aname, ",")
        path = Right(aname, Len(aname) - x)
        aname = Left(aname, x - 1)
        lstApps.ListItems.Add i, , aname
        lstApps.ListItems(i).ListSubItems.Add , , path
    End If
Next
End Sub

Private Sub cmdFolder_Click()
txtFav.Text = BrowseForFolder(Me.hwnd, "NC3允许你选择一个空的文件夹用来存放书签，也可以将书签设置为IE收藏夹路径，从而和IE收藏夹并存")
End Sub

Private Sub Form_Load()
tabOption_Click
Frame(0).Height = tabOption.Height - 420
Frame(0).Width = tabOption.Width - 240
Frame(1).Top = Frame(0).Top
Frame(1).Left = Frame(0).Left
Frame(1).Height = tabOption.Height - 420
Frame(1).Width = tabOption.Width - 240

LoadApps
txtHome.Text = Loadini(OpFile, "navigator", "home", "http://fsslinux.51.net")
txtFav.Text = Loadini(OpFile, "navigator", "favorite", "c:\windows\favorites")
chkAllowNewWindow.Value = Loadini(OpFile, "window", "enablenewwindow", 1)
chkXPStyle.Value = Loadini(OpFile, "window", "xp", 0)
chkShowSideBar.Value = Loadini(OpFile, "window", "sidebar", 1)
End Sub

Public Sub tabOption_Click()
Dim i As Integer
For i = 0 To tabOption.Tabs.Count - 1
    If i = tabOption.SelectedItem.Index - 1 Then
        Frame(i).Visible = True
    Else
        Frame(i).Visible = False
    End If
Next
End Sub
