VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DC4F4966-71B5-44BA-A7DE-759B67636012}#1.0#0"; "cPopMenu.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "NetCaravel III"
   ClientHeight    =   7950
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10770
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   10770
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picSplitter 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   7800
      ScaleHeight     =   3735
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox imgSplitter 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   7560
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3135
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   3120
      Width           =   60
   End
   Begin VB.DirListBox favDir 
      Height          =   300
      Left            =   9000
      TabIndex        =   58
      Top             =   6480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.FileListBox favFile 
      Height          =   270
      Left            =   9000
      Pattern         =   "*.url"
      TabIndex        =   57
      Top             =   6720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Index           =   2
      Left            =   120
      ScaleHeight     =   5505
      ScaleWidth      =   3225
      TabIndex        =   15
      Top             =   1440
      Width           =   3255
      Begin VB.CommandButton cmdBookmarkGO 
         Caption         =   "GO"
         Height          =   255
         Left            =   2400
         TabIndex        =   62
         Top             =   430
         Width           =   615
      End
      Begin VB.TextBox txtBookmarkURL 
         Enabled         =   0   'False
         Height          =   270
         Left            =   0
         TabIndex        =   61
         Top             =   400
         Width           =   2415
      End
      Begin VB.CommandButton cmdBookmarkRefresh 
         Caption         =   "刷新书签"
         Height          =   375
         Left            =   1920
         TabIndex        =   60
         Top             =   0
         Width           =   975
      End
      Begin MSComctlLib.TreeView TreeBookmark 
         Height          =   4575
         Left            =   0
         TabIndex        =   59
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   8070
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         Style           =   7
         HotTracking     =   -1  'True
         SingleSel       =   -1  'True
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBookmarkDEL 
         Caption         =   "删除书签"
         Height          =   375
         Left            =   960
         TabIndex        =   56
         Top             =   0
         Width           =   975
      End
      Begin VB.CommandButton cmdBookmarkAdd 
         Caption         =   "添加书签"
         Height          =   375
         Left            =   0
         TabIndex        =   55
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   1
      Left            =   1560
      ScaleHeight     =   5625
      ScaleWidth      =   3585
      TabIndex        =   14
      Top             =   1440
      Width           =   3615
      Begin VB.CommandButton cmdGoogleGuide 
         Caption         =   "搜索帮助"
         Height          =   375
         Left            =   1440
         TabIndex        =   47
         Tag             =   "http://www.google.com/intl/zh-CN/help.html"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdGoogleForum 
         Caption         =   "Google大全"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Tag             =   "http://www.google.com/intl/zh-CN/about.html"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdGoogleADV 
         Caption         =   "高级搜索"
         Height          =   375
         Left            =   1440
         TabIndex        =   54
         Tag             =   "http://www.google.com/advanced_search?hl=zh-CN"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdGoogleEN 
         Caption         =   "英文Google"
         Height          =   375
         Left            =   120
         TabIndex        =   53
         Tag             =   "http://www.google.com/en"
         Top             =   480
         Width           =   1335
      End
      Begin VB.CommandButton cmdGooglePI 
         Caption         =   "使用偏好"
         Height          =   375
         Left            =   1440
         TabIndex        =   52
         Tag             =   "http://groups.google.com/preferences?q=%E4%B8%AD%E6%96%87&hl=zh-CN&lr=&ie=UTF-8"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdGoogleLanguage 
         Caption         =   "语言工具"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Tag             =   "http://www.google.com/language_tools?hl=zh-CN"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.TextBox txtGoogleKey 
         Height          =   270
         Left            =   120
         TabIndex        =   50
         Text            =   "关键词"
         Top             =   120
         Width           =   1695
      End
      Begin VB.CommandButton cmdGoogleGO 
         Caption         =   "开始搜索"
         Height          =   375
         Left            =   1920
         TabIndex        =   49
         Top             =   60
         Width           =   975
      End
      Begin VB.Frame FrameGoogle1 
         Caption         =   "搜索类型(Google及其附加引擎)"
         Height          =   1335
         Left            =   120
         TabIndex        =   43
         Top             =   960
         Width           =   3375
         Begin VB.OptionButton OptionGoogle 
            Caption         =   "搜索新闻组群"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   45
            Tag             =   "http://groups.google.com/groups?q="
            Top             =   960
            Width           =   1455
         End
         Begin VB.OptionButton OptionGoogle 
            Caption         =   "搜索图像"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   44
            Tag             =   "http://images.google.com/images?q="
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptionGoogle 
            Caption         =   "搜索网页"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Tag             =   "http://www.google.com/search?q="
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame FrameGoogle2 
         Caption         =   "缩小搜索范围"
         Height          =   1575
         Left            =   120
         TabIndex        =   37
         Top             =   2400
         Width           =   3375
         Begin VB.CheckBox chkGoogleADV 
            Caption         =   "使用缩小搜索范围功能"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   2295
         End
         Begin VB.CheckBox chkGoogleADV 
            Caption         =   "在指定站点搜索"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   41
            Top             =   600
            Width           =   1575
         End
         Begin VB.CheckBox chkGoogleADV 
            Caption         =   "去除多余的关键词"
            Height          =   180
            Index           =   2
            Left            =   120
            TabIndex        =   40
            Top             =   960
            Width           =   1815
         End
         Begin VB.TextBox txtGoogleSite 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1920
            TabIndex        =   39
            Text            =   "指定站点，不包含http://"
            ToolTipText     =   "请输入站点的地址，不包含http://"
            Top             =   570
            Width           =   1215
         End
         Begin VB.TextBox txtGoogleKeyEX 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   2040
            TabIndex        =   38
            Text            =   "关键词"
            Top             =   930
            Width           =   1095
         End
      End
      Begin VB.Image ImageLogo 
         Height          =   795
         Index           =   1
         Left            =   240
         MouseIcon       =   "frmMain.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":0BD4
         Tag             =   "http://www.google.com"
         ToolTipText     =   "点击访问Google主页"
         Top             =   4920
         Width           =   1920
      End
   End
   Begin VB.PictureBox PicBox 
      Appearance      =   0  'Flat
      FillColor       =   &H80000008&
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   0
      Left            =   3840
      ScaleHeight     =   5625
      ScaleWidth      =   3345
      TabIndex        =   6
      Top             =   1440
      Width           =   3375
      Begin VB.Frame FrameBaidu2 
         Caption         =   "百度高级搜索功能"
         Height          =   1575
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   3135
         Begin VB.TextBox txtBaiduKeyEX 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1920
            TabIndex        =   26
            Text            =   "关键词"
            Top             =   1170
            Width           =   1095
         End
         Begin VB.TextBox txtBaiduSite 
            Appearance      =   0  'Flat
            Height          =   270
            Left            =   1800
            TabIndex        =   27
            Text            =   "指定站点，不包含http://"
            ToolTipText     =   "请输入站点的地址，不包含http://"
            Top             =   450
            Width           =   1215
         End
         Begin VB.CheckBox chkBaiduADV 
            Caption         =   "去除多余的关键词"
            Height          =   180
            Index           =   4
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Width           =   1815
         End
         Begin VB.CheckBox chkBaiduADV 
            Caption         =   "仅搜索网页地址"
            Height          =   180
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
         Begin VB.CheckBox chkBaiduADV 
            Caption         =   "仅搜索网页标题"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Width           =   1575
         End
         Begin VB.CheckBox chkBaiduADV 
            Caption         =   "在指定站点搜索"
            Height          =   180
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1575
         End
         Begin VB.CheckBox chkBaiduADV 
            Caption         =   "使用高级功能"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame FrameBaidu1 
         Caption         =   "搜索类型(百度及其附加引擎)"
         Height          =   1815
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   3135
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索城市天气预报"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   24
            Tag             =   "http://www.t7online.com/cgi-bin/suchen?LANG=cn&PRG=citybild&ORT="
            ToolTipText     =   "中国天气在线"
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索信息快递"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   23
            Tag             =   "http://www1.baidu.com/wstsearch?tn=baiduwstui&ct=83886080&lm=-1&bs=%CF%F1%BB%A8%D2%BB%D1%F9&word="
            Top             =   1200
            Width           =   1575
         End
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索歌词"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   22
            Tag             =   "http://mp3.baidu.com/wstsearch?tn=baidump3lyric&ct=150994944&rn=10&word="
            Top             =   960
            Width           =   1095
         End
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索Flash"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Tag             =   "http://flash.baidu.com/wstsearch?tn=flash&ct=33554432&word="
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索MP3"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   20
            Tag             =   "http://mp3.baidu.com/wstsearch?tn=baidump3&ct=134217728&rn=&word="
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptionBaidu 
            Caption         =   "搜索网页"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Tag             =   "http://www1.baidu.com/baidu?word="
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdBaiduGuide 
         Caption         =   "搜索指南"
         Height          =   375
         Left            =   1440
         TabIndex        =   36
         Tag             =   "http://www.baidu.com/search/jiqiao.html#27"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaiduForum 
         Caption         =   "搜索援助中心"
         Height          =   375
         Left            =   120
         TabIndex        =   35
         Tag             =   "http://forum.baidu.com/cgi-bin/forum/board_show.cgi?id=1&age=60"
         Top             =   4440
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaiduPost 
         Caption         =   "邮编区号查询"
         Height          =   375
         Left            =   1440
         TabIndex        =   34
         Tag             =   "http://lib.wyu.edu.cn/post/search.asp"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaiduIP 
         Caption         =   "IP地址查询"
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Tag             =   "http://www.yofoo.com/ipq/default.asp"
         Top             =   4080
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaiduGO 
         Caption         =   "开始搜索"
         Height          =   375
         Left            =   1920
         TabIndex        =   17
         Top             =   60
         Width           =   975
      End
      Begin VB.TextBox txtBaiduKey 
         Height          =   270
         Left            =   120
         TabIndex        =   16
         Text            =   "关键词"
         Top             =   120
         Width           =   1695
      End
      Begin VB.Image ImageLogo 
         Height          =   795
         Index           =   0
         Left            =   240
         MouseIcon       =   "frmMain.frx":1AE5
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":1DEF
         Tag             =   "http://www.baidu.com"
         ToolTipText     =   "点击访问百度主页"
         Top             =   4920
         Width           =   1920
      End
   End
   Begin MSComDlg.CommonDialog CDLG 
      Left            =   8880
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ImageList imgSmallIcon 
      Left            =   8280
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC2
            Key             =   "tabbed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F13
            Key             =   "url"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3345
            Key             =   "www"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":378E
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39F6
            Key             =   "fav"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Progress 
      Height          =   135
      Left            =   120
      TabIndex        =   13
      Top             =   7800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin SHDocVwCtl.WebBrowser Web 
      Height          =   1095
      Index           =   0
      Left            =   6360
      TabIndex        =   5
      Tag             =   "0"
      Top             =   1440
      Width           =   1095
      ExtentX         =   1931
      ExtentY         =   1931
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   7695
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3519
            MinWidth        =   3528
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11298
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   4
            TextSave        =   "0:33"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1746
            MinWidth        =   4
            TextSave        =   "2022-1-17"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo comboSogua 
      Height          =   315
      Left            =   9120
      TabIndex        =   10
      ToolTipText     =   "输入关键词，按回车后即可在SoGua搜索引擎搜索"
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   16576
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "SoGua搜索(MP3)"
   End
   Begin MSComctlLib.ImageCombo comboBaidu 
      Height          =   315
      Left            =   6000
      TabIndex        =   9
      ToolTipText     =   "输入关键词，按回车后即可在百度搜索引擎搜索"
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   255
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "百度搜索(中)"
   End
   Begin MSComctlLib.ImageCombo comboURL 
      Height          =   315
      Left            =   0
      TabIndex        =   7
      ToolTipText     =   "输入地址后，按下回车即可访问Web"
      Top             =   600
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   8388608
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TabStrip Tabbed 
      Height          =   1935
      Left            =   6240
      TabIndex        =   3
      Top             =   960
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   3413
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "网页1"
            Object.Tag             =   "0"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picBackground 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   45
      Left            =   7080
      Picture         =   "frmMain.frx":3DF0
      ScaleHeight     =   45
      ScaleWidth      =   6750
      TabIndex        =   0
      Top             =   5280
      Visible         =   0   'False
      Width           =   6750
   End
   Begin cPopMenu.PopMenu PopMenu 
      Left            =   9480
      Top             =   2280
      _ExtentX        =   1058
      _ExtentY        =   1058
      HighlightCheckedItems=   0   'False
      TickIconIndex   =   0
      BorderColor     =   12582912
      HForeColor      =   16777215
      ShadowXPHighlight=   0   'False
      ShadowXPHighlightTopMenu=   0   'False
   End
   Begin MSComctlLib.ImageList imgToolbarHot 
      Left            =   9480
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48C0
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4CB5
            Key             =   "back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F28
            Key             =   "bookmark"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5341
            Key             =   "option"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5771
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5B9C
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5CFD
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6146
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65D7
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A63
            Key             =   "print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6EE1
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7386
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":75F8
            Key             =   "home"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7A3A
            Key             =   "help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7E8C
            Key             =   "misc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82E1
            Key             =   "apps"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":870A
            Key             =   "reload"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B60
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8FA3
            Key             =   "close"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9394
            Key             =   "new"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgToolbar 
      Left            =   8880
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   22
      ImageHeight     =   22
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   20
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":97DC
            Key             =   "add"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9921
            Key             =   "back"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9BA7
            Key             =   "bookmark"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D6B
            Key             =   "option"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F4F
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1EC
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A34D
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A558
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A860
            Key             =   "open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AB66
            Key             =   "print"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE6C
            Key             =   "save"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B182
            Key             =   "forward"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B40B
            Key             =   "home"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B6DF
            Key             =   "help"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9BC
            Key             =   "misc"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC8E
            Key             =   "apps"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BF54
            Key             =   "reload"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C165
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C373
            Key             =   "close"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C524
            Key             =   "new"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo comboGoogle 
      Height          =   315
      Left            =   7560
      TabIndex        =   8
      ToolTipText     =   "输入关键词，按回车后即可在google搜索引擎搜索"
      Top             =   600
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      ForeColor       =   32768
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "Google搜索(英)"
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10770
      _ExtentX        =   18997
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.TabStrip SideBar 
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "百度(中文)"
            Object.ToolTipText     =   "百度高级搜索功能"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Google(国际)"
            Object.ToolTipText     =   "Google高级搜索功能"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "书签"
            Object.ToolTipText     =   "使用书签功能，记录经常访问的Web地址"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFileNew 
         Caption         =   "新建标签(&T)"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "另存为...(&S)"
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "关闭(&C)"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePageSetup 
         Caption         =   "页面设置(&I)"
      End
      Begin VB.Menu mnuFilePrintPriview 
         Caption         =   "打印预览(&O)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
      End
      Begin VB.Menu mnuFileLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePro 
         Caption         =   "页面属性(&W)"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)     "
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditCut 
         Caption         =   "剪切(&X)"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "复制(&C)"
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "粘贴(&V)"
      End
      Begin VB.Menu mnuEditLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelall 
         Caption         =   "全选(&A)"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "查找(&F)"
      End
   End
   Begin VB.Menu mnuNavigate 
      Caption         =   "导航(&N)"
      Begin VB.Menu mnuNavigateBack 
         Caption         =   "后退(&B)"
      End
      Begin VB.Menu mnuNavigateForward 
         Caption         =   "前进(&F)"
      End
      Begin VB.Menu mnuNavigateLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuNavigateReload 
         Caption         =   "刷新(&R)"
      End
      Begin VB.Menu mnuNavigateStop 
         Caption         =   "停止(&S)"
      End
      Begin VB.Menu mnuNavigateGoHome 
         Caption         =   "主页(&H)"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuToolsNewWindow 
         Caption         =   "允许网页打开新标签(&A)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuToolsSetHome 
         Caption         =   "将目前的地址设置为主页(&H)"
      End
      Begin VB.Menu mnuToolsDefaultBrowser 
         Caption         =   "使用默认浏览器浏览此页面(&D)"
      End
      Begin VB.Menu mnuToolsApplinks 
         Caption         =   "外接程序(&L)"
         Begin VB.Menu mnuToolsApps 
            Caption         =   "-"
            Index           =   0
         End
      End
      Begin VB.Menu mnuToolsLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsPing 
         Caption         =   "Ping(&P)"
      End
      Begin VB.Menu mnuToolsNetstat 
         Caption         =   "Netstat(&N)"
      End
      Begin VB.Menu mnuToolsIpconfig 
         Caption         =   "Ipconfig(&I)"
      End
      Begin VB.Menu mnuToolsLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsOption 
         Caption         =   "选项(&O)"
      End
      Begin VB.Menu mnuToolsInternet 
         Caption         =   "Internet Explorer选项(&E)"
      End
   End
   Begin VB.Menu mnuBookmark 
      Caption         =   "书签(&B)"
      Begin VB.Menu mnuBookmarkAdd 
         Caption         =   "添加(&D)"
      End
      Begin VB.Menu mnuBookmarkShow 
         Caption         =   "显示书签(&B)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelphelp 
         Caption         =   "帮助(&H)"
      End
      Begin VB.Menu mnuHelpGoFSS 
         Caption         =   "访问主页(&G)"
      End
      Begin VB.Menu mnuHelpSendMail 
         Caption         =   "发送邮件给作者(&S)"
      End
      Begin VB.Menu mnuHelpLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ebMoving As Boolean
Const sglSplitLimit = 500
Public sIndex As Integer
Public urlHome As String
Public webNumber As Integer
Public BookmarkPath As String

Public Sub LoadApplications()
On Error Resume Next
Dim i As Integer, x As Integer, path As String, aname As String
For i = 1 To 40
    Unload mnuToolsApps(i)
    aname = Loadini(OpFile, "apps", "app" & Str(i))
    If aname <> "" Then
        Load mnuToolsApps(i)
        x = InStr(1, aname, ",")
        path = Right(aname, Len(aname) - x)
        aname = Left(aname, x - 1)
        mnuToolsApps(i).Caption = aname
        mnuToolsApps(i).Tag = path
    End If
Next
mnuToolsApps(0).Visible = False
End Sub

Public Sub LoadBookmarks()
On Error Resume Next

If BookmarkPath = "" Then
    TreeBookmark.Nodes.Add , , "root", "尚未选择正确的书签路径，请在选项中设置", "fav"
    Exit Sub
End If

favDir.path = BookmarkPath
favDir.Refresh
favFile.Refresh
TreeBookmark.Nodes(1).Expanded = False
TreeBookmark.Nodes.Clear
TreeBookmark.Nodes.Add , , "root", "书签", "fav"

Dim i As Integer, i2 As Integer, path As String, FolderName As String, FileName As String

For i = 0 To favDir.ListCount - 1
    path = favDir.List(i)
    FolderName = Right(path, Len(path) - Len(BookmarkPath))
    
    TreeBookmark.Nodes.Add "root", tvwChild, FolderName, FolderName, "folder"
    favFile.path = path
    
    For i2 = 0 To favFile.ListCount - 1
        FileName = Left(favFile.List(i2), Len(favFile.List(i2)) - 4)
        TreeBookmark.Nodes.Add FolderName, tvwChild, favFile.path & "\" & favFile.List(i2), FileName, "www"
        If Loadini(favFile.path & "\" & favFile.List(i2), "InternetShortcut", "URL") <> "" Then _
             TreeBookmark.Nodes(favFile.path & "\" & favFile.List(i2)).Tag = _
             Loadini(favFile.path & "\" & favFile.List(i2), "InternetShortcut", "URL")
    Next
Next

favFile.path = BookmarkPath

For i2 = 0 To favFile.ListCount - 1
        FileName = Left(favFile.List(i2), Len(favFile.List(i2)) - 4)
        TreeBookmark.Nodes.Add "root", tvwChild, favFile.path & "\" & favFile.List(i2), FileName, "www"
        If Loadini(favFile.path & "\" & favFile.List(i2), "InternetShortcut", "URL") <> "" Then _
             TreeBookmark.Nodes(favFile.path & "\" & favFile.List(i2)).Tag = _
             Loadini(favFile.path & "\" & favFile.List(i2), "InternetShortcut", "URL")
Next

TreeBookmark.Nodes(1).Expanded = True

End Sub

Public Sub Delbookmark(path As String)
On Error Resume Next
If Right(path, 4) = ".url" Or Right(path, 4) = ".URL" Then
Kill path
Else
path = BookmarkPath & path
Dim yes As Integer
yes = MsgBox("你要删除整个目录的书签，确认继续？", vbYesNo + 64, "问题")
    If yes = vbYes Then
        Open "c:\del.bat" For Output As #1
        Print #1, "@echo off"
        Print #1, "rmdir " & Chr(34) & path & Chr(34) & " /q/s"
        Print #1, "del c:\del.bat"
        Close #1
        Shell "c:\del.bat", vbHide
    End If

End If
LoadBookmarks

End Sub

Public Sub NewWeb(Optional URL As String)
    i = Val(Tabbed.Tabs(Tabbed.Tabs.Count).Tag) + 1
    Load Web(i)
    Web(i).Visible = False
    Web(i).Left = Tabbed.Left + 30
    Web(i).Top = Tabbed.Top + 300
    Web(i).Width = Tabbed.Width - 30
    Web(i).Height = Tabbed.Height - 300
    Web(i).ZOrder 0
    Web(i).Tag = i
    If URL = "" Then
        Web(i).Navigate2 "about:blank"
    Else
        Web(i).Navigate2 URL
    End If
    webNumber = webNumber + 1
    sTab = Tabbed.Tabs.Count + 1
    Tabbed.Tabs.Add , , "网页" & sTab, "tabbed"
    Tabbed.Tabs(sTab).Tag = i
    Tabbed.Tabs(sTab).Selected = True
    Tabbed_Click
    SizeControls (imgSplitter.Left)
End Sub

Sub SizeControls(x As Single)
On Error Resume Next
Dim i As Integer

If tbrMain.Buttons("misc").Value = tbrPressed Then

    If x < 1500 Then x = 2500
    If x > (Me.ScaleWidth - 2500) Then x = Me.ScaleWidth - 2500
    imgSplitter.Top = Me.tbrMain.Height + Me.comboURL.Height
    imgSplitter.Left = x
    imgSplitter.Height = Me.ScaleHeight - Me.tbrMain.Height - Me.StatusBar.Height - Me.comboURL.Height
    picSplitter.Height = imgSplitter.Height
    
    SideBar.Visible = True
    SideBar.Top = imgSplitter.Top
    SideBar.Left = 0
    SideBar.Height = imgSplitter.Height
    SideBar.Width = x
    SideBar_Click

    For i = 0 To PicBox.Count - 1
        PicBox(i).Left = 30
        PicBox(i).Top = SideBar.Top + 300
        PicBox(i).Width = SideBar.Width - 30
        PicBox(i).Height = SideBar.Height - 300
    Next

Else
    x = 0
    imgSplitter.Top = Me.tbrMain.Height + Me.comboURL.Height
    imgSplitter.Left = x
    imgSplitter.Height = Me.ScaleHeight - Me.tbrMain.Height - Me.StatusBar.Height - Me.comboURL.Height
    picSplitter.Height = imgSplitter.Height

    SideBar.Visible = False
    For i = 0 To PicBox.Count - 1
        PicBox(i).Visible = False
    Next

End If

comboURL.Top = tbrMain.Height
comboGoogle.Top = comboURL.Top
comboBaidu.Top = comboURL.Top
comboSogua.Top = comboURL.Top
comboURL.Width = Me.ScaleWidth - comboGoogle.Width * 3
comboBaidu.Left = comboURL.Width
comboGoogle.Left = comboBaidu.Left + comboBaidu.Width
comboSogua.Left = comboGoogle.Left + comboGoogle.Width


Tabbed.Top = tbrMain.Height + comboURL.Height
Tabbed.Height = imgSplitter.Height
Tabbed.Left = imgSplitter.Left + 60
Tabbed.Width = Me.ScaleWidth - x - 60

For i = 0 To Web.Count - 1
    Web(i).Left = Tabbed.Left + 30
    Web(i).Top = Tabbed.Top + 300
    Web(i).Width = Tabbed.Width - 30
    Web(i).Height = Tabbed.Height - 300
Next

Progress.Left = 30
Progress.Top = Tabbed.Top + Tabbed.Height + 60
Progress.Height = StatusBar.Height - 60
Progress.Width = 1960

FrameBaidu1.Width = PicBox(0).ScaleWidth - FrameBaidu1.Left - 120
FrameBaidu2.Width = FrameBaidu1.Width
FrameGoogle1.Width = FrameBaidu1.Width
FrameGoogle2.Width = FrameBaidu1.Width

Me.txtBaiduKey.Width = FrameBaidu1.Width - Me.txtBaiduKey.Left - Me.cmdBaiduGO.Width - 120
Me.cmdBaiduGO.Left = Me.txtBaiduKey.Width + Me.txtBaiduKey.Left + 120
Me.txtBaiduSite.Width = Me.FrameBaidu2.Width - Me.chkBaiduADV(1).Width - Me.chkBaiduADV(1).Left - 240
Me.txtBaiduKeyEX.Width = Me.FrameBaidu2.Width - Me.chkBaiduADV(4).Width - Me.chkBaiduADV(4).Left - 120

Me.txtGoogleKey.Width = txtBaiduKey.Width
Me.cmdGoogleGO.Left = cmdBaiduGO.Left
Me.txtGoogleSite.Width = FrameGoogle2.Width - Me.chkGoogleADV(1).Width - Me.chkGoogleADV(1).Left - 360
Me.txtGoogleKeyEX.Width = FrameGoogle2.Width - Me.chkGoogleADV(2).Width - Me.chkGoogleADV(2).Left - 240

Me.TreeBookmark.Height = Me.PicBox(2).ScaleHeight - TreeBookmark.Top
Me.TreeBookmark.Width = Me.PicBox(2).ScaleWidth
Me.txtBookmarkURL.Width = Me.PicBox(2).ScaleWidth - cmdBookmarkGO.Width - 120
Me.cmdBookmarkGO.Left = Me.txtBookmarkURL.Width + 120
End Sub


Private Sub chkBaiduADV_Click(Index As Integer)
Dim i As Integer
Select Case Index
    Case 0
        If chkBaiduADV(0).Value = 1 Then
            For i = 1 To 4
                chkBaiduADV(i).Enabled = True
            Next
        Else
            For i = 1 To 4
                chkBaiduADV(i).Enabled = False
            Next
        End If
    Case 1
        Me.txtBaiduSite.Enabled = chkBaiduADV(1).Value
    Case 2
        If chkBaiduADV(2).Value = 1 Then
            Me.chkBaiduADV(3).Value = 0
        End If
        
    Case 3
        If chkBaiduADV(3).Value = 1 Then
            Me.chkBaiduADV(2).Value = 0
        End If
    Case 4
        Me.txtBaiduKeyEX.Enabled = chkBaiduADV(4).Value
End Select
End Sub

Private Sub chkGoogleADV_Click(Index As Integer)
Select Case Index
    Case 0
        If chkGoogleADV(Index).Value = 1 Then
            chkGoogleADV(1).Enabled = True
            chkGoogleADV(2).Enabled = True
        Else
            chkGoogleADV(1).Enabled = False
            chkGoogleADV(2).Enabled = False
        End If
    Case 1
        Me.txtGoogleSite.Enabled = chkGoogleADV(1).Value
    Case 2
        Me.txtGoogleKeyEX.Enabled = chkGoogleADV(2).Value
End Select
End Sub

Private Sub cmdBaiduForum_Click()
Web(sIndex).Navigate2 cmdBaiduForum.Tag
End Sub

Private Sub cmdBaiduGO_Click()
Dim URL As String, i As Integer, KEY As String
For i = 0 To 5
    If Me.OptionBaidu(i).Value = True Then
        URL = Me.OptionBaidu(i).Tag
    End If
Next

KEY = Me.txtBaiduKey.Text

If chkBaiduADV(0).Enabled = True And chkBaiduADV(0).Value = 1 Then
    If chkBaiduADV(1).Value = 1 Then KEY = KEY & " site:" & Me.txtBaiduSite
    If chkBaiduADV(2).Value = 1 Then KEY = "intitle:" & KEY
    If chkBaiduADV(3).Value = 1 Then KEY = "inurl:" & KEY
    If chkBaiduADV(4).Value = 1 Then KEY = KEY & " -" & Me.txtBaiduKeyEX
End If

URL = URL & KEY

Web(sIndex).Navigate2 URL

End Sub

Private Sub cmdBaiduGuide_Click()
Web(sIndex).Navigate2 cmdBaiduGuide.Tag
End Sub

Private Sub cmdBaiduIP_Click()
Web(sIndex).Navigate2 cmdBaiduIP.Tag
End Sub

Private Sub cmdBaiduPost_Click()
Web(sIndex).Navigate2 cmdBaiduPost.Tag
End Sub

Private Sub cmdBookmarkAdd_Click()
On Error GoTo Errors
CDLG.DialogTitle = "保存书签"
CDLG.FileName = BookmarkPath & Me.Tabbed.SelectedItem.Caption & ".url"
CDLG.Filter = "Internet快捷方式|*.url|"
CDLG.ShowSave
Saveini CDLG.FileName, "InternetShortcut", "URL", Web(sIndex).LocationURL
LoadBookmarks

Exit Sub

Errors:

End Sub

Private Sub cmdBookmarkDEL_Click()
Delbookmark TreeBookmark.Nodes(TreeBookmark.SelectedItem.Index).KEY
LoadBookmarks
End Sub

Private Sub cmdBookmarkGO_Click()
Web(sIndex).Navigate2 txtBookmarkURL.Text
End Sub

Private Sub cmdBookmarkRefresh_Click()
LoadBookmarks
End Sub

Private Sub cmdGoogleADV_Click()
Web(sIndex).Navigate2 cmdGoogleADV.Tag
End Sub

Private Sub cmdGoogleEN_Click()
Web(sIndex).Navigate2 cmdGoogleEN.Tag
End Sub

Private Sub cmdGoogleForum_Click()
Web(sIndex).Navigate2 cmdGoogleForum.Tag
End Sub

Private Sub cmdGoogleGO_Click()
Dim URL As String, i As Integer, KEY As String
For i = 0 To 2
    If OptionGoogle(i).Value = True Then
        URL = OptionGoogle(i).Tag
    End If
Next

KEY = txtGoogleKey.Text

If chkGoogleADV(0).Enabled = True And chkGoogleADV(0).Value = 1 Then
    If chkGoogleADV(1).Value = 1 Then KEY = KEY & " site:" & Me.txtGoogleSite.Text
    If chkGoogleADV(2).Value = 1 Then KEY = KEY & " -" & Me.txtGoogleKeyEX.Text
End If

URL = URL & KEY

Web(sIndex).Navigate2 URL

End Sub

Private Sub cmdGoogleGuide_Click()
Web(sIndex).Navigate2 cmdGoogleGuide.Tag
End Sub

Private Sub cmdGoogleLanguage_Click()
Web(sIndex).Navigate2 cmdGoogleLanguage.Tag
End Sub

Private Sub cmdGooglePI_Click()
Web(sIndex).Navigate2 cmdGooglePI.Tag
End Sub

Private Sub comboBaidu_GotFocus()
comboBaidu.SelStart = 0
comboBaidu.SelLength = LenB(comboBaidu.Text)
End Sub

Private Sub comboBaidu_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
comboBaidu.ComboItems.Add 1, , comboBaidu.Text
Web(sIndex).Navigate2 "http://www1.baidu.com/baidu?word=" & comboBaidu.Text
End If
End Sub

Private Sub comboGoogle_GotFocus()
comboGoogle.SelStart = 0
comboGoogle.SelLength = LenB(comboGoogle.Text)
End Sub

Private Sub comboGoogle_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
comboGoogle.ComboItems.Add 1, , comboGoogle.Text
Web(sIndex).Navigate2 "http://www.google.com/search?hl=en&ie=UTF-8&oe=UTF-8&q=" & comboGoogle.Text
End If
End Sub

Private Sub comboSogua_GotFocus()
comboSogua.SelStart = 0
comboSogua.SelLength = LenB(comboSogua.Text)
End Sub

Private Sub comboSogua_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
comboSogua.ComboItems.Add 1, , comboSogua.Text
Web(sIndex).Navigate2 "http://search.sogua.com/search/search.asp?key=" & comboSogua.Text
End If
End Sub

Private Sub comboURL_GotFocus()
comboURL.SelStart = 0
comboURL.SelLength = LenB(comboURL.Text)
End Sub

Private Sub comboURL_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
comboURL.ComboItems.Add 1, , comboURL.Text, "url"
comboURL.ComboItems(1).Selected = True
Web(sIndex).Navigate2 comboURL.Text
End If
End Sub

Private Sub Form_Load()

On Error Resume Next

    Dim i As Integer, msg As Integer

If Loadini(OpFile, "window", "firstrun", 1) = 1 Then
    msg = MsgBox("这是你第一运行nc3，你使用的是nt内核的系统(win2000,winxp等)吗？请正确回答，否则会引起菜单无法显示", vbYesNo + 64, "提问")
    If msg = vbYes Then
        Saveini OpFile, "window", "xp", 1
    Else
        Saveini OpFile, "window", "xp", 0
    End If
    Saveini OpFile, "window", "firstrun", 0
End If
    
    If Loadini(OpFile, "window", "xp", 0) = 1 Then
    PopMenu.SubClassMenu Me
    PopMenu.HighlightStyle = cspHighlightXP
    PopMenu.ShadowXPHighlight = True
    Set PopMenu.BackgroundPicture = picBackground.Picture
    PopMenu.HighlightForeColor = &H0&
    End If

'创建工具栏

            Set tbrMain.ImageList = imgToolbar
            Set tbrMain.HotImageList = imgToolbarHot
            
            With tbrMain.Buttons
            Set Y = .Add(, "new", , , "new")
                Y.ToolTipText = "打开新的Web浏览窗口"
            Set Y = .Add(, "close", , , "close")
                Y.ToolTipText = "关闭当前窗口"
            Set Y = .Add(, "open", , , "open")
                Y.ToolTipText = "打开一个本地网页"
            Set Y = .Add(, "save", , , "save")
                Y.ToolTipText = "另存网页"
            Set Y = .Add
                Y.Style = tbrPlaceholder
            Set Y = .Add(, "back", , , "back")
                Y.ToolTipText = "后退"
            Set Y = .Add(, "forward", , , "forward")
                Y.ToolTipText = "前进"
            Set Y = .Add
                Y.Style = tbrPlaceholder
            Set Y = .Add(, "reload", , , "reload")
                Y.ToolTipText = "重载"
            Set Y = .Add(, "stop", , , "stop")
                Y.ToolTipText = "停止"
            Set Y = .Add(, "gohome", , , "home")
                Y.ToolTipText = "主页"
            Set Y = .Add
                Y.Style = tbrPlaceholder
            Set Y = .Add(, "print", , , "print")
                Y.ToolTipText = "打印"
            Set Y = .Add(, "preview", , , "preview")
                Y.ToolTipText = "打印预览"
            Set Y = .Add
                Y.Style = tbrPlaceholder
            Set Y = .Add(, "cut", , , "cut")
                Y.ToolTipText = "剪切"
            Set Y = .Add(, "copy", , , "copy")
                Y.ToolTipText = "复制"
            Set Y = .Add(, "paste", , , "paste")
                Y.ToolTipText = "粘贴"
            Set Y = .Add
                Y.Style = tbrPlaceholder
            Set Y = .Add(, "misc", , , "misc")
                Y.ToolTipText = "边条"
                Y.Style = tbrCheck
            Set Y = .Add(, "bookmark", , , "bookmark")
                Y.ToolTipText = "书签"
            Set Y = .Add(, "add", , , "add")
                Y.ToolTipText = "加入书签"
            Set Y = .Add(, "apps", , , "apps")
                Y.ToolTipText = "外接"
            Set Y = .Add(, "option", , , "option")
                Y.ToolTipText = "选项"
            Set Y = .Add
                Y.Style = tbrPlaceholder

            End With
            
            
'初始化窗口&&读入设置
Me.WindowState = Loadini(OpFile, "window", "state", 0)
Me.Left = Loadini(OpFile, "window", "x", 0)
Me.Top = Loadini(OpFile, "window", "y", 0)
Me.Height = Loadini(OpFile, "window", "height", 6400)
Me.Width = Loadini(OpFile, "window", "width", 8000)
tbrMain.Buttons("misc").Value = Loadini(OpFile, "window", "sidebar", 1)
imgSplitter.Left = Loadini(OpFile, "window", "splitter", 2500)
urlHome = Loadini(OpFile, "navigator", "home", "http://fsslinux.51.net")
Me.mnuToolsNewWindow.Checked = Loadini(OpFile, "navigator", "enablenewwindow", 1)

Set Tabbed.ImageList = Me.imgSmallIcon
Tabbed.Tabs(1).Image = "tabbed"
Set comboURL.ImageList = Me.imgSmallIcon
SideBar_Click

For i = 0 To 4
    chkBaiduADV_Click (i)
Next
For i = 0 To 2
    chkGoogleADV_Click (i)
Next

SizeControls (imgSplitter.Left)
webNumber = 1
Web(0).Navigate2 urlHome

comboBaidu.ComboItems.Add , , "百度搜索(中)"
comboGoogle.ComboItems.Add , , "Google搜索(英)"
comboSogua.ComboItems.Add , , "Sogua搜索(mp3)"

Set TreeBookmark.ImageList = Me.imgSmallIcon
BookmarkPath = Loadini(OpFile, "navigator", "favorite", App.path & "\bookmarks") & "\"
Me.favDir.path = BookmarkPath
Me.favFile.path = BookmarkPath
LoadBookmarks
LoadApplications

End Sub

Private Sub Form_Resize()
SizeControls (imgSplitter.Left)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Saveini OpFile, "window", "x", Me.Left
Saveini OpFile, "window", "y", Me.Top
Saveini OpFile, "window", "width", Me.Width
Saveini OpFile, "window", "height", Me.Height
Saveini OpFile, "window", "state", Me.WindowState
Saveini OpFile, "window", "splitter", Me.imgSplitter.Left
Saveini OpFile, "window", "sidebar", tbrMain.Buttons("misc").Value
Saveini OpFile, "navigator", "enablenewwindow", mnuToolsNewWindow.Checked
End
End Sub

Private Sub ImageLogo_Click(Index As Integer)
Web(sIndex).Navigate2 ImageLogo(Index).Tag
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    ebMoving = True
End Sub

Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim sglPos As Single


    If ebMoving Then
       sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If

End Sub

Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    ebMoving = False

End Sub

Private Sub mnuBookmarkAdd_Click()
cmdBookmarkAdd_Click
End Sub

Private Sub mnuBookmarkShow_Click()
If Me.tbrMain.Buttons("misc").Value = tbrUnpressed Then
Me.tbrMain.Buttons("misc").Value = tbrPressed
SizeControls (2500)
End If
Me.SideBar.Tabs(3).Selected = True
SideBar_Click
End Sub

Private Sub mnuEditCopy_Click()
Web(sIndex).ExecWB OLECMDID_COPY, 0, 0, 0
End Sub

Private Sub mnuEditCut_Click()
Web(sIndex).ExecWB OLECMDID_CUT, 0, 0, 0
End Sub

Private Sub mnuEditFind_Click()
    SendKeys "^(f)"
    Web(sIndex).SetFocus
End Sub

Private Sub mnuEditPaste_Click()
Web(sIndex).ExecWB OLECMDID_PASTE, 0, 0, 0
End Sub

Private Sub mnuEditSelall_Click()
Web(sIndex).ExecWB OLECMDID_SELECTALL, 0, 0, 0
End Sub

Private Sub mnuFileClose_Click()
    If Tabbed.Tabs.Count = 1 Then
        MsgBox "不能删除所有的标签框，nc3要求至少一个标签框显示"
    Else
        Unload Web(sIndex)
        Tabbed.Tabs.Remove Tabbed.SelectedItem.Index
        Tabbed_Click
    End If
End Sub

Private Sub mnuFileExit_Click()
 End
End Sub

Private Sub mnuFileNew_Click()
    NewWeb
End Sub

Private Sub mnuFileOpen_Click()

On Error GoTo Errors
    CDLG.DialogTitle = "打开网页"
    CDLG.Filter = "超文本文档(*.htm;*.html)|*.htm;*.html|图形文件(*.bmp;*.gif;*.jpg;*.png)|*.bmp;*.gif;*.jpg;*.png;*.jpeg|文本文件(*.txt)|*.txt|Flash动画(*.swf)|*.swf|所有文件(*.*)|*.*|"
    CDLG.ShowOpen
    Web(sIndex).Navigate2 CDLG.FileName
    
Exit Sub


Errors:
End Sub

Private Sub mnuFilePageSetup_Click()
    Web(sIndex).ExecWB OLECMDID_PAGESETUP, 0, 0, 0
End Sub

Private Sub mnuFilePrint_Click()
    Web(sIndex).ExecWB OLECMDID_PRINT, 0, 0, 0
End Sub

Private Sub mnuFilePrintPriview_Click()
    Web(sIndex).ExecWB OLECMDID_PRINTPREVIEW, 0, 0, 0
End Sub

Private Sub mnuFilePro_Click()
Web(sIndex).ExecWB OLECMDID_PROPERTIES, 0, 0, 0
End Sub

Private Sub mnuFileSave_Click()
    Web(sIndex).ExecWB OLECMDID_SAVEAS, 0, 0, 0
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show vbModal, Me
End Sub

Private Sub mnuHelpGoFSS_Click()
Web(sIndex).Navigate2 "http://fsslinux.51.net"
End Sub

Private Sub mnuHelphelp_Click()
Web(sIndex).Navigate2 App.path & "\help\index.htm"
End Sub

Private Sub mnuHelpSendMail_Click()
     ShellExecute hwnd, "Open", "mailto:zxh@yeah.net", 0, 0, 0
End Sub

Private Sub mnuNavigateBack_Click()
On Error Resume Next
Web(sIndex).GoBack
End Sub

Private Sub mnuNavigateForward_Click()
On Error Resume Next
Web(sIndex).GoForward
End Sub

Private Sub mnuNavigateGoHome_Click()
On Error Resume Next
Web(sIndex).Navigate2 urlHome
End Sub

Private Sub mnuNavigateReload_Click()
On Error Resume Next
Web(sIndex).Refresh2
End Sub

Private Sub mnuNavigateStop_Click()
On Error Resume Next
Web(sIndex).Stop
End Sub

Private Sub mnuToolsApps_Click(Index As Integer)
Shell mnuToolsApps(Index).Tag, vbNormalFocus
End Sub

Private Sub mnuToolsDefaultBrowser_Click()
     ShellExecute hwnd, "Open", Me.Web(sIndex).LocationURL, 0, 0, 0
End Sub

Private Sub mnuToolsInternet_Click()
Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
End Sub

Private Sub mnuToolsIpconfig_Click()
Open "c:\ipconfig.bat" For Output As #1
Print #1, "@echo off"
Print #1, "call ipconfig > c:\nc.txt"
Print #1, "call notepad c:\nc.txt"
Print #1, "call del c:\nc.txt"
Print #1, "call del c:\ipconfig.bat"
Close #1
Shell "c:\ipconfig.bat", vbNormalFocus
End Sub

Private Sub mnuToolsNetstat_Click()
Open "c:\netstat.bat" For Output As #1
Print #1, "@echo off"
Print #1, "echo 根据网络连接状况，Netstat可能会耗费几十秒到数分钟的时间，请耐心等待。"
Print #1, "echo 如果不想继续等待，可以按CTRL+C终止"
Print #1, "call netstat > c:\nc.txt"
Print #1, "call notepad c:\nc.txt"
Print #1, "call del c:\nc.txt"
Print #1, "call del c:\netstat.bat"
Close #1
Shell "c:\netstat.bat", vbNormalFocus
End Sub

Private Sub mnuToolsNewWindow_Click()
If mnuToolsNewWindow.Checked = True Then
    mnuToolsNewWindow.Checked = False
Else
    mnuToolsNewWindow.Checked = True
End If
End Sub

Private Sub mnuToolsOption_Click()
frmOptions.Show vbModal, Me
End Sub

Private Sub mnuToolsPing_Click()
Dim URL As String
URL = InputBox("清输入ping的地址", "Ping", "127.0.0.1")
Open "c:\ping.bat" For Output As #1
Print #1, "@echo off"
Print #1, "echo 根据网络状况，ping一个地址需要数秒钟的时间左右，请耐心等待"
Print #1, "call ping " & URL & "> c:\nc.txt"
Print #1, "call notepad c:\nc.txt"
Print #1, "call del c:\nc.txt"
Print #1, "call del c:\ping.bat"
Close #1
Shell "c:\ping.bat", vbNormalFocus
End Sub

Private Sub mnuToolsSetHome_Click()
Saveini OpFile, "navigator", "home", Web(sIndex).LocationURL
End Sub

Private Sub OptionBaidu_Click(Index As Integer)
Dim i As Integer
If Index <> 0 Then
    For i = 0 To 4
        chkBaiduADV(i).Enabled = False
    Next
Else
    For i = 0 To 4
        chkBaiduADV(i).Enabled = True
    Next
End If
End Sub

Private Sub PopMenu_ItemHighlight(ItemNumber As Long, bEnabled As Boolean, bSeparator As Boolean)
StatusBar.Panels(3).Text = "菜单：" & PopMenu.Caption(ItemNumber)
End Sub

Private Sub SideBar_Click()
Dim i As Integer
For i = 0 To PicBox.Count - 1
    If SideBar.SelectedItem.Index = i + 1 Then
        PicBox(i).Visible = True
    Else
        PicBox(i).Visible = False
    End If
Next
End Sub

Private Sub Tabbed_Click()
On Error Resume Next
Dim i As Integer

For i = 0 To webNumber
    If i = Val(Tabbed.SelectedItem.Tag) Then
        Web(i).ZOrder 0
        Web(i).Visible = True
    Else
        Web(i).Visible = False
    End If
Next

sIndex = Val(Tabbed.SelectedItem.Tag)
Me.Caption = Tabbed.SelectedItem.Caption & " - NetCaravel III"
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim i As Integer, sTab As Integer
Select Case Button.KEY
    Case "back"
        Web(sIndex).GoBack
    Case "forward"
        Web(sIndex).GoForward
    Case "stop"
        Web(sIndex).Stop
    Case "reload"
        Web(sIndex).Refresh2
    Case "new"
        Me.NewWeb (urlHome)
    Case "close"
        mnuFileClose_Click
    Case "open"
        mnuFileOpen_Click
    Case "save"
        mnuFileSave_Click
    Case "misc"
        If Button.Value = tbrUnpressed Then
            imgSplitter.Enabled = False
            SizeControls (3000)
        Else
            imgSplitter.Enabled = True
            SizeControls (3000)
        End If
    Case "cut"
        mnuEditCut_Click
    Case "copy"
        mnuEditCopy_Click
    Case "paste"
        mnuEditPaste_Click
    Case "preview"
        mnuFilePrintPriview_Click
    Case "print"
        mnuFilePrint_Click
    Case "bookmark"
        mnuBookmarkShow_Click
    Case "add"
        mnuBookmarkAdd_Click
    Case "option"
        mnuToolsOption_Click
    Case "apps"
        frmOptions.tabOption.Tabs(2).Selected = True
        frmOptions.tabOption_Click
        frmOptions.Show
End Select

End Sub

Private Sub TreeBookmark_NodeClick(ByVal Node As MSComctlLib.Node)
If Node.Tag <> "" Then
Me.txtBookmarkURL.Text = Node.Tag
End If
End Sub

Private Sub Web_DocumentComplete(Index As Integer, ByVal pDisp As Object, URL As Variant)
comboURL.ComboItems.Add 1, , URL, "url"
comboURL.ComboItems(1).Selected = True
SizeControls (imgSplitter.Left)
End Sub

Private Sub Web_NewWindow2(Index As Integer, ppDisp As Object, Cancel As Boolean)
If mnuToolsNewWindow.Checked = True Then
    NewWeb
    Set ppDisp = Web(sIndex).Object
Else
    Cancel = True
End If
End Sub

Private Sub Web_ProgressChange(Index As Integer, ByVal Progress As Long, ByVal ProgressMax As Long)
On Error Resume Next
Me.Progress.Max = ProgressMax
Me.Progress.Value = Progress
End Sub

Private Sub Web_StatusTextChange(Index As Integer, ByVal Text As String)
Dim i As Integer
For i = 1 To Tabbed.Tabs.Count
    If Web(Index).Tag = Tabbed.Tabs(i).Tag Then
    StatusBar.Panels(3).Text = Text
    End If
Next
End Sub

Private Sub Web_TitleChange(Index As Integer, ByVal Text As String)
Dim i As Integer
For i = 1 To Tabbed.Tabs.Count
    If Web(Index).Tag = Tabbed.Tabs(i).Tag Then
    Tabbed.Tabs(i).Caption = Text
    StatusBar.Panels(2).Text = Text
    End If
Next
Me.Caption = Text & " - NetCaravel III"
End Sub
