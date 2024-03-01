VERSION 5.00
Begin VB.Form LoginScreen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "管理员登陆"
   ClientHeight    =   1110
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   3735
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   249
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      Caption         =   "登陆"
      Height          =   372
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   3732
   End
   Begin VB.TextBox Text2 
      Height          =   264
      IMEMode         =   3  'DISABLE
      Left            =   0
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   360
      Width           =   3732
   End
   Begin VB.TextBox Text1 
      Height          =   264
      IMEMode         =   3  'DISABLE
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3732
   End
End
Attribute VB_Name = "LoginScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Private Sub Command1_Click()
    If Text1.Text = "lkzy" And Text2.Text = "67081752" Or Text1.Text = "lyzk" And Text2.Text = "dqh2529082" Then
        If Text1.Text = "lkzy" Then denglu = "开发者选项 当前登陆账号:裂空之云" Else denglu = "开发者选项 当前登陆账号:裂云之空"
        DeveloperMode.Move 0, Screen.Height - DeveloperMode.Height
        DeveloperMode.Show 0
        MainForm.Top = DeveloperMode.Top - MainForm.Height '非全屏状态独有
        If MainForm.v = 4 Then '非全屏状态独有
            moshi4.Top = DeveloperMode.Top - MainForm.Height '非全屏状态独有
        ElseIf MainForm.v = 6 Then '非全屏状态独有
            moshi6.Top = DeveloperMode.Top - MainForm.Height '非全屏状态独有
        End If '非全屏状态独有
        DeveloperMode.Caption = denglu
        Unload LoginScreen
        MainForm.f8 = True
    Else
        MsgBox "账号或密码错误", vbExclamation + vbSystemModal, "登录失败"
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
    End If
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    LoginScreen.Show
    Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Command1_Click
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Command1_Click
End Sub
