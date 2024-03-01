VERSION 5.00
Begin VB.Form moshi6 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   Caption         =   "2048(6*6)"
   ClientHeight    =   9375
   ClientLeft      =   10035
   ClientTop       =   4545
   ClientWidth     =   9375
   Icon            =   "Form4.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   625
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   625
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   35
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   34
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   33
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   32
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   31
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   30
      Left            =   0
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   29
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   28
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   27
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   26
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   25
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   24
      Left            =   0
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   23
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   22
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   21
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   20
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   19
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   18
      Left            =   0
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   17
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   16
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   15
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   14
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   13
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   12
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   11
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   10
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   9
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   8
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   7
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   6
      Left            =   0
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   5
      Left            =   7800
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   4
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   3
      Left            =   4680
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   2
      Left            =   3120
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   1
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
   Begin VB.Image Image1 
      Height          =   1572
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1572
   End
End
Attribute VB_Name = "moshi6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
    
Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        x1 = x: y1 = y
    ElseIf Button = 2 Then
        PopupMenu MainForm.gameshezhi
    End If
End Sub

Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x3 As Single, Y3 As Single)
    If f5 = False Or MainForm.f7 = True Or MainForm.Timer3.Enabled = True Then Exit Sub
    Call Delay(0.5)
    If Button = 1 Then
        If Abs(Y3 - y1) > Abs(x3 - x1) Then
            If Y3 - y1 > 0 Then
                Call MainForm.MainProgress(83)
            ElseIf Y3 - y1 < 0 Then
                Call MainForm.MainProgress(87)
            End If
        ElseIf Abs(Y3 - y1) < Abs(x3 - x1) Then
            If x3 - x1 > 0 Then
                Call MainForm.MainProgress(68)
            ElseIf x3 - x1 < 0 Then
                Call MainForm.MainProgress(65)
            End If
        End If
    End If
End Sub

Private Sub Form_keypress(keyascii As Integer)
    If MainForm.Timer3.Enabled = False And MainForm.f7 = False Then Call MainForm.MainProgress(keyascii)
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        tcbc = MsgBox("您的游戏窗口正处于打开状态,是否保存当前进度？", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "退出之一")
        If tcbc = vbYes Then Call MainForm.bc
        MainForm.Timer2.Enabled = False
        DeveloperMode.List1.Clear
        DeveloperMode.List2.Clear
        Unload moshi6
    End If
End Sub

Sub tmdc()
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, MainForm.tmd3, LWA_ALPHA
End Sub
