VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于网际快帆"
   ClientHeight    =   5415
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5970
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3737.529
   ScaleMode       =   0  'User
   ScaleWidth      =   5606.139
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   4320
      Width           =   1500
   End
   Begin VB.Label lblInfo 
      Caption         =   "https://www.github.com/vimac"
      Height          =   255
      Left            =   120
      MousePointer    =   2  'Cross
      TabIndex        =   3
      Top             =   5040
      Width           =   5655
   End
   Begin VB.Label lblVersion 
      Caption         =   "Label1"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   4440
      Width           =   2055
   End
   Begin VB.Label lblTitle 
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   3960
      Width           =   2535
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
Set Me.Picture = frmSplash.Picture
lblVersion.Caption = "版本 " & App.Major & "." & App.Minor & "." & App.Revision
lblTitle.Caption = "Net Caravel II"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmSplash
End Sub

Private Sub lblInfo_Click()
     ShellExecute hwnd, "Open", "https://www.github.com/vimac", 0, 0, 0
End Sub
