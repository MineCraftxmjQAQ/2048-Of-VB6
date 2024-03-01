VERSION 5.00
Begin VB.Form DeveloperMode 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "开发者模式"
   ClientHeight    =   4935
   ClientLeft      =   5790
   ClientTop       =   3705
   ClientWidth     =   11175
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   745
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command17 
      Caption         =   "产生一致数"
      Height          =   495
      Left            =   3000
      TabIndex        =   28
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command16 
      Caption         =   "读取存档"
      Height          =   492
      Left            =   7320
      TabIndex        =   27
      Top             =   1560
      Width           =   1332
   End
   Begin VB.CommandButton Command15 
      Caption         =   "保存存档"
      Height          =   492
      Left            =   5880
      TabIndex        =   26
      Top             =   1560
      Width           =   1332
   End
   Begin VB.CommandButton Command14 
      Caption         =   "玩死检测"
      Height          =   492
      Left            =   4440
      TabIndex        =   25
      Top             =   2160
      Width           =   1332
   End
   Begin VB.CommandButton Command13 
      Caption         =   "游戏说明"
      Height          =   492
      Left            =   5880
      TabIndex        =   24
      Top             =   2160
      Width           =   1332
   End
   Begin VB.CommandButton Command12 
      Caption         =   "2048更新日志"
      Height          =   492
      Left            =   7320
      TabIndex        =   23
      Top             =   2160
      Width           =   1332
   End
   Begin VB.HScrollBar HScroll5 
      Height          =   252
      LargeChange     =   10
      Left            =   2400
      Max             =   255
      TabIndex        =   22
      Top             =   4560
      Width           =   8652
   End
   Begin VB.CommandButton Command11 
      Caption         =   "下一步"
      Height          =   492
      Left            =   3000
      TabIndex        =   20
      Top             =   1560
      Width           =   1332
   End
   Begin VB.CommandButton Command10 
      Caption         =   "上一步"
      Height          =   492
      Left            =   1560
      TabIndex        =   19
      Top             =   1560
      Width           =   1332
   End
   Begin VB.HScrollBar HScroll4 
      Height          =   252
      LargeChange     =   10
      Left            =   2400
      Max             =   255
      TabIndex        =   18
      Top             =   4200
      Width           =   8652
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   252
      LargeChange     =   10
      Left            =   2400
      Max             =   255
      TabIndex        =   17
      Top             =   3840
      Width           =   8652
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   252
      LargeChange     =   10
      Left            =   2400
      Max             =   255
      TabIndex        =   14
      Top             =   3480
      Width           =   8652
   End
   Begin VB.CommandButton Command5 
      Caption         =   "退出登陆"
      Height          =   492
      Left            =   9720
      TabIndex        =   12
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CommandButton Command4 
      Caption         =   "滑动开关"
      Height          =   492
      Left            =   4440
      TabIndex        =   11
      Top             =   1560
      Width           =   1332
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   252
      LargeChange     =   5
      Left            =   2400
      Max             =   101
      Min             =   1
      TabIndex        =   9
      Top             =   3120
      Value           =   1
      Width           =   8652
   End
   Begin VB.CommandButton Command9 
      Caption         =   "顺序切歌"
      Height          =   492
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   1332
   End
   Begin VB.CommandButton Command8 
      Caption         =   "显示音乐控件"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1332
   End
   Begin VB.CommandButton Command7 
      Caption         =   "切换游戏模式"
      Height          =   495
      Left            =   8760
      TabIndex        =   6
      Tag             =   "两大模式"
      Top             =   1560
      Width           =   1332
   End
   Begin VB.ListBox List2 
      Height          =   1320
      Left            =   7080
      TabIndex        =   5
      Top             =   120
      Width           =   3972
   End
   Begin VB.CommandButton Command6 
      Caption         =   "切换难度等级"
      Height          =   492
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   1332
   End
   Begin VB.ListBox List1 
      Height          =   1320
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   5412
   End
   Begin VB.CommandButton Command3 
      Caption         =   "显示6*6方格"
      Height          =   492
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1332
   End
   Begin VB.CommandButton Command2 
      Caption         =   "显示多人对战"
      Height          =   492
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      Caption         =   "显示4*4方格"
      Height          =   492
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1332
   End
   Begin VB.Label Label5 
      Height          =   252
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   2292
   End
   Begin VB.Label Label4 
      Height          =   252
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   2292
   End
   Begin VB.Label Label3 
      Height          =   252
      Left            =   120
      TabIndex        =   15
      Top             =   3840
      Width           =   2292
   End
   Begin VB.Label Label2 
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   3480
      Width           =   2292
   End
   Begin VB.Label Label1 
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   3120
      Width           =   2292
   End
End
Attribute VB_Name = "DeveloperMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2

Private Sub Command10_Click()
    If MainForm.moveclick.Checked = True Then
        MainForm.moveclick.Checked = False
        Call MainForm.huitui
        MainForm.moveclick.Checked = True
    Else
        Call MainForm.huitui
    End If
End Sub

Private Sub Command11_Click()
    If MainForm.moveclick.Checked = True Then
        MainForm.moveclick.Checked = False
        Call MainForm.qianjin
        MainForm.moveclick.Checked = True
    Else
        Call MainForm.qianjin
    End If
End Sub

Private Sub Command12_Click()
     UpdateForm.Show 1
End Sub

Private Sub Command13_Click()
    about.Show 1
End Sub

Private Sub Command14_Click()
    If MainForm.moveclick.Checked = True Then
        MainForm.moveclick.Checked = False
        If MainForm.Timer3.Enabled = False And MainForm.f7 = False Then Call MainForm.ws
        MainForm.moveclick.Checked = True
    Else
        If MainForm.Timer3.Enabled = False And MainForm.f7 = False Then Call MainForm.ws
    End If
End Sub

Private Sub Command15_Click()
    Call MainForm.bc
End Sub

Private Sub Command16_Click()
    Call MainForm.loadgame_Click
End Sub

Private Sub Command17_Click()
    If MainForm.moveclick.Checked = True Then
        MainForm.moveclick.Checked = False
        If MainForm.Timer3.Enabled = False And MainForm.f7 = False Then Call MainForm.yz
        MainForm.moveclick.Checked = True
    Else
        If MainForm.Timer3.Enabled = False And MainForm.f7 = False Then Call MainForm.yz
    End If
End Sub

Private Sub Command2_Click()
    Call MainForm.qingchu
    Call MainForm.Image2_DblClick
End Sub

Private Sub Command4_Click()
    Call MainForm.moveclick_Click
    If MainForm.moveclick.Checked = True Then Command4.Caption = "关闭滑动开关" Else Command4.Caption = "开启滑动开关"
End Sub

Private Sub Command5_Click()
    MainForm.f8 = False
    Unload DeveloperMode
    LoginScreen.Show
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Command4.Caption = "开启滑动开关"
    DeveloperMode.Caption = denglu
    HScroll1.Value = MainForm.ylgl
    HScroll2.Value = MainForm.tmd1
    HScroll3.Value = MainForm.tmd2
    HScroll4.Value = MainForm.tmd3
    HScroll5.Value = MainForm.tmd4
End Sub

Private Sub Command1_Click()
    If MainForm.Timer3.Enabled = True Then Exit Sub
    Call MainForm.qingchu
    Call MainForm.Image1_DblClick
End Sub

Private Sub Command3_Click()
    If MainForm.Timer3.Enabled = True Then Exit Sub
    Call MainForm.qingchu
    Call MainForm.Image3_DblClick
End Sub


Private Sub Command6_Click()
    MainForm.nd = MainForm.nd + 1
    If MainForm.nd = 4 Then MainForm.nd = 1
    If MainForm.nd = 1 Then
        MainForm.level1.Enabled = False
        MainForm.level2.Enabled = True
        MainForm.level3.Enabled = True
    ElseIf MainForm.nd = 2 Then
        MainForm.level1.Enabled = True
        MainForm.level2.Enabled = False
        MainForm.level3.Enabled = True
    ElseIf MainForm.nd = 3 Then
        MainForm.level1.Enabled = True
        MainForm.level2.Enabled = True
        MainForm.level3.Enabled = False
    End If
    MainForm.dh1 = 0
    Call MainForm.sc
End Sub

Private Sub Command7_Click()
    MainForm.ms = MainForm.ms + (-1) ^ (MainForm.ms + 1)
    MainForm.ordinary.Enabled = Not MainForm.ordinary.Enabled
    MainForm.youle.Enabled = Not MainForm.youle.Enabled
    MainForm.prev.Enabled = Not MainForm.prev.Enabled
    Command10.Enabled = Not Command10.Enabled
    Command11.Enabled = Not Command11.Enabled
    MainForm.dh1 = 0
    Call MainForm.sc
End Sub

Private Sub Command8_Click()
    MainForm.Picture1.Visible = Not MainForm.Picture1.Visible
    If MainForm.Picture1.Visible = False Then
        Command8.Caption = "显示音乐控件"
    ElseIf MainForm.Picture1.Visible = True Then
        Command8.Caption = "隐藏音乐控件"
    End If
End Sub

Sub Command9_Click()
    Call MainForm.sxqg
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload DeveloperMode
    End If
End Sub

Sub tmdd()
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, MainForm.tmd4, LWA_ALPHA
End Sub

Private Sub HScroll1_Change()
    Label1.Caption = "当前游乐概率:" + CStr(101 - HScroll1.Value) + "%"
    MainForm.ylgl = HScroll1.Value
End Sub

Private Sub HScroll2_Change()
    Label2.Caption = "当前主交互透明度:" + Mid(CStr(HScroll2.Value / 255 * 100), 1, 5) + "%"
    MainForm.tmd1 = HScroll2.Value
    If HScroll2.Value / 255 * 100 < 20 Then
        MsgBox "为了操作体验和显示效果，主交互界面的透明度不得低于20%", 0 + vbExclamation + vbSystemModal, "提示"
        HScroll2.Value = 51
        MainForm.tmd1 = HScroll2.Value
    End If
    Call MainForm.tmda
End Sub

Private Sub HScroll3_Change()
    Label3.Caption = "当前4*4透明度:" + Mid(CStr(HScroll3.Value / 255 * 100), 1, 5) + "%"
    MainForm.tmd2 = HScroll3.Value
    Call moshi4.tmdb
End Sub

Private Sub HScroll4_Change()
    Label4.Caption = "当前6*6透明度:" + Mid(CStr(HScroll4.Value / 255 * 100), 1, 5) + "%"
    MainForm.tmd3 = HScroll4.Value
    Call moshi6.tmdc
End Sub

Private Sub HScroll5_Change()
    Label5.Caption = "当前开发者透明度:" + Mid(CStr(HScroll5.Value / 255 * 100), 1, 5) + "%"
    MainForm.tmd4 = HScroll5.Value
        If HScroll5.Value / 255 * 100 < 60 Then
        MsgBox "为了操作体验和显示效果，开发者模式界面的透明度不得低于60%", 0 + vbExclamation + vbSystemModal, "提示"
        HScroll5.Value = 153
        MainForm.tmd4 = HScroll5.Value
    End If
    Call tmdd
End Sub
