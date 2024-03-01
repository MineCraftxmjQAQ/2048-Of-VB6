VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainForm 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7665
   ClientLeft      =   12045
   ClientTop       =   2550
   ClientWidth     =   10680
   ControlBox      =   0   'False
   FillColor       =   &H00FF80FF&
   ForeColor       =   &H8000000B&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   511
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer4 
      Left            =   6480
      Top             =   6360
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7440
      Top             =   5760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Left            =   6960
      Top             =   5760
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7920
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6480
      Top             =   5760
   End
   Begin VB.PictureBox Picture1 
      Height          =   1515
      Left            =   2040
      ScaleHeight     =   1455
      ScaleWidth      =   2190
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   2244
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   1500
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2220
         URL             =   ""
         rate            =   1
         balance         =   0
         currentPosition =   0
         defaultFrame    =   ""
         playCount       =   1
         autoStart       =   -1  'True
         currentMarker   =   0
         invokeURLs      =   -1  'True
         baseURL         =   ""
         volume          =   50
         mute            =   0   'False
         uiMode          =   "full"
         stretchToFit    =   0   'False
         windowlessVideo =   0   'False
         enabled         =   -1  'True
         enableContextMenu=   -1  'True
         fullScreen      =   0   'False
         SAMIStyle       =   ""
         SAMILang        =   ""
         SAMIFilename    =   ""
         captioningID    =   ""
         enableErrorDialogs=   0   'False
         _cx             =   3916
         _cy             =   2646
      End
   End
   Begin VB.Image Image1 
      Height          =   2004
      Left            =   0
      Picture         =   "Form1.frx":038A
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2004
   End
   Begin VB.Image Image4 
      Height          =   5628
      Left            =   0
      Picture         =   "Form1.frx":4537D
      Stretch         =   -1  'True
      Top             =   0
      Width           =   10620
   End
   Begin VB.Image Image3 
      Height          =   2004
      Left            =   8640
      Picture         =   "Form1.frx":2C1AAC
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2004
   End
   Begin VB.Image Image2 
      Height          =   2004
      Left            =   4320
      Picture         =   "Form1.frx":2FDDD9
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   2004
   End
   Begin VB.Menu gameshezhi 
      Caption         =   "游戏设置"
      Visible         =   0   'False
      Begin VB.Menu prev 
         Caption         =   "上一步"
      End
      Begin VB.Menu nexts 
         Caption         =   "下一步"
      End
      Begin VB.Menu moveclick 
         Caption         =   "滑动开关"
         Checked         =   -1  'True
      End
      Begin VB.Menu baocun 
         Caption         =   "保存存档"
      End
      Begin VB.Menu loadgame 
         Caption         =   "读取存档"
      End
      Begin VB.Menu moshi 
         Caption         =   "模式"
         Begin VB.Menu ordinary 
            Caption         =   "经典模式"
         End
         Begin VB.Menu youle 
            Caption         =   "游乐模式"
         End
      End
      Begin VB.Menu level 
         Caption         =   "难度"
         Begin VB.Menu level1 
            Caption         =   "普通"
         End
         Begin VB.Menu level2 
            Caption         =   "困难"
         End
         Begin VB.Menu level3 
            Caption         =   "地狱"
         End
      End
      Begin VB.Menu pcsong 
         Caption         =   "顺序切歌"
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助"
      Visible         =   0   'False
      Begin VB.Menu youxishuoming 
         Caption         =   "游戏说明"
      End
      Begin VB.Menu update 
         Caption         =   "2048更新日志"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public first As Integer
Public js As Integer
Public last As Integer
Public nd As Integer
Public ms As Integer
Public v As Integer
Public q As String
Public yy As Integer
Public ylgl As Integer
Public dh1 As Integer

Public tmd1 As Byte
Public tmd2 As Byte
Public tmd3 As Byte
Public tmd4 As Byte

Public f7 As Boolean
Public f8 As Boolean

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim t As Integer
Dim time As Integer
Dim u As Integer
Dim r As Integer
Dim m As Integer
Dim n As Integer

Dim nc As Integer
Dim ncmax As Integer
Dim mov1 As Integer
Dim count1 As Integer

Dim tp As Integer
Dim tt As Integer
Dim max As Integer
Dim tc As Integer
Dim yl As Integer
Dim mszt As Integer
Dim mszt1 As Integer
Dim hdzb1 As Integer

Dim dh2 As Single
Dim dh3 As Single

Dim wz1 As Integer
Dim ylwz As Integer
Dim ylsz As Integer

Dim a(0 To 36) As Integer
Dim b(1 To 36) As Integer
Dim c(1 To 36) As Integer
Dim d(1 To 41) As Integer
Dim e(1 To 37, 1 To 2) As Integer
Dim f(1 To 37) As Integer
Dim g(1 To 37, 1 To 2) As Integer
Dim ct As Form

Dim s As String
Dim xrsc As String
Dim returntime As String
Dim mx(0 To 4) As String
Dim xam(0 To 4) As String
Dim pets(0 To 4) As String

Dim f1 As Boolean
Dim f2 As Boolean
Dim f3 As Boolean
Dim f4 As Boolean
Dim f6 As Boolean
Sub Form_Load()
    moveclick.Checked = False
    max = 0: first = 0: js = 0: last = 0: nd = 1: ms = 1: tp = 0: yy = 1: ylgl = 96: tmd1 = 255: tmd2 = 255: tmd3 = 255: tmd4 = 255
    Erase a
    Erase b
    Erase c
    Erase d
    Erase e
    Erase f
    Erase g
    f5 = False
    musiclist(1) = "\music\MineCraft\C418 - Minecraft.mp3"
    musiclist(2) = "\music\MineCraft\C418 - Beginning.mp3"
    musiclist(3) = "\music\MineCraft\C418 - Dry Hands.mp3"
    musiclist(4) = "\music\MineCraft\C418 - Danny.mp3"
    musiclist(5) = "\music\MineCraft\C418 - Sweden.mp3"
    musiclist(6) = "\music\MineCraft\C418 - Door.mp3"
    musiclist(7) = "\music\MineCraft\C418 - Droopy likes your face.mp3"
    musiclist(8) = "\music\MineCraft\C418 - Cat.mp3"
    musiclist(9) = "\music\MineCraft\C418 - équinoxe.mp3"
    musiclist(10) = "\music\MineCraft\C418 - Excuse.mp3"
    musiclist(11) = "\music\MineCraft\C418 - Haggstrom.mp3"
    musiclist(12) = "\music\MineCraft\C418 - Key.mp3"
    musiclist(13) = "\music\MineCraft\C418 - Living Mice.mp3"
    musiclist(14) = "\music\MineCraft\C418 - Mice on Venus.mp3"
    musiclist(15) = "\music\MineCraft\C418 - Moog City.mp3"
    musiclist(16) = "\music\MineCraft\C418 - Oxygène.mp3"
    musiclist(17) = "\music\MineCraft\C418 - Subwoofer Lullaby.mp3"
    musiclist(18) = "\music\MineCraft\C418 - Dog.mp3"
    musiclist(19) = "\music\MineCraft\C418 - Thirteen.mp3"
    musiclist(20) = "\music\MineCraft\An Ordinary Day - Kumi Tanioka.mp3"
    musiclist(21) = "\music\MineCraft\Ancestry - Lena Raine.mp3"
    musiclist(22) = "\music\MineCraft\Comforting Memories - Kumi Tanioka.mp3"
    musiclist(23) = "\music\MineCraft\Floating Dream - Kumi Tanioka.mp3"
    musiclist(24) = "\music\MineCraft\Infinite Amethyst - Lena Raine.mp3"
    musiclist(25) = "\music\MineCraft\Left to Bloom - Lena Raine.mp3"
    musiclist(26) = "\music\MineCraft\One More Day - Lena Raine.mp3"
    musiclist(27) = "\music\MineCraft\Otherside - Lena Raine.mp3"
    musiclist(28) = "\music\MineCraft\Stand Tall - Lena Raine.mp3"
    musiclist(29) = "\music\MineCraft\Wending - Lena Raine.mp3"
    WindowsMediaPlayer1.URL = App.Path & "\music\MineCraft\C418 - Minecraft.mp3"
End Sub

Sub ws()
    If MainForm.v = 4 Then
        For i = 1 To 16
            a(i) = i + 1
        Next i
    ElseIf MainForm.v = 6 Then
        For i = 1 To 16
            a(i) = i + 1
            a(16 + i) = a(i)
        Next i
        a(33) = 1
        a(34) = 2
        a(35) = 3
        a(36) = 4
    End If
    Call qingchu
    Call sc
End Sub

Sub yz()
    If MainForm.v = 4 Then
        For i = 1 To 16
            a(i) = 1
        Next i
    ElseIf MainForm.v = 6 Then
        For i = 1 To 36
            a(i) = 1
        Next i
    End If
    Call qingchu
    Call sc
End Sub

Function levelzfc(level As Integer) As String
    If level = 1 Then
        levelzfc = "普通"
    ElseIf level = 2 Then
        levelzfc = "困难"
    ElseIf level = 3 Then
        levelzfc = "地狱"
    End If
End Function

Sub huitui() 'excel屏蔽
    If js = first Or js = 1 Then MsgBox "不要得寸进尺", vbExclamation + vbSystemModal, "无法回退": Exit Sub
    For i = 1 To v * v
        a(i) = xlSheet.Cells(i + 1, js - 1)
    Next i
    js = js - 1
    Call qingchu
    Call sc
End Sub

Sub qianjin() 'excel屏蔽
    If js = last Then MsgBox "你已经到达了世界的尽头", vbExclamation + vbSystemModal, "无法前进": Exit Sub
    For i = 1 To v * v
        a(i) = xlSheet.Cells(i + 1, js + 1)
    Next i
    js = js + 1
    Call qingchu
    Call sc
End Sub

Sub music()
    yy = yy + 1
    If yy = 30 Then yy = 1
End Sub

Sub sxqg()
    Call music
    WindowsMediaPlayer1.URL = App.Path & musiclist(MainForm.yy)
    Call jiance
End Sub

Sub qingchu()
    For i = 1 To 36
        b(i) = 97
    Next i
    Erase e
    Erase f
    Erase g
End Sub

Sub youxishuoming_Click()
    about.Show 1
End Sub

Sub Image1_DblClick() '全屏与非全屏不同的过程
    Call caidan
    v = 4
    Call qingchu
    Call tongyong
    Unload moshi6
    Unload duorenmoshi
    DeveloperMode.List2.Left = 472
    DeveloperMode.List1.Width = 360.8
    DeveloperMode.List2.Width = 264.8
    moshi4.Top = 0
    moshi4.Left = Screen.Width - moshi4.Width
    moshi4.Show 0
    Call jiance
End Sub

Sub Image2_DblClick()
    Call caidan
    v = 5
    Call qingchu
    Call tongyong
    Unload moshi4
    Unload moshi6
    DeveloperMode.List1.Width = 312.8
    DeveloperMode.List2.Width = 312.8
    DeveloperMode.List2.Left = 424.8
    duorenmoshi.Show 0
    Call jiance
End Sub

Sub Image3_DblClick() '全屏与非全屏不同的过程
    Call caidan
    v = 6
    Call qingchu
    Call tongyong
    Unload moshi4
    Unload duorenmoshi
    DeveloperMode.List2.Left = 472
    DeveloperMode.List1.Width = 360.8
    DeveloperMode.List2.Width = 264.8
    moshi6.Top = 0
    moshi6.Left = Screen.Width - moshi6.Width
    moshi6.Show 0
    Call jiance
End Sub

Sub Image4_DblClick()
    If f8 = False Then
        LoginScreen.Show 0
    Else
        DeveloperMode.Show 0
        DeveloperMode.Move 0, Screen.Height - DeveloperMode.Height
    End If
End Sub

Sub caidan()
    ms = 1: nd = 1
    ordinary.Enabled = False
    youle.Enabled = True
    level1.Enabled = False
    level2.Enabled = True
    level3.Enabled = True
    DeveloperMode.Command10.Enabled = True
    DeveloperMode.Command11.Enabled = True
End Sub

Sub level1a()
    r = MsgBox("您确定要修改难度为普通？此操作需要重新开始游戏。", vbOKCancel + vbExclamation + vbSystemModal, "切换难度")
    If r = 1 Then
        nd = 1
        Call tongyong
        level1.Enabled = False
        level2.Enabled = True
        level3.Enabled = True
    End If
End Sub

Sub level2a()
    r = MsgBox("您确定要修改难度为困难？此操作需要重新开始游戏。", vbOKCancel + vbExclamation + vbSystemModal, "切换难度")
    If r = 1 Then
        nd = 2
        Call tongyong
        level1.Enabled = True
        level2.Enabled = False
        level3.Enabled = True
    End If
End Sub

Sub level3a()
    r = MsgBox("您确定要修改难度为地狱？此操作需要重新开始游戏。", vbOKCancel + vbExclamation + vbSystemModal, "切换难度")
    If r = 1 Then
        nd = 3
        Call tongyong
        level1.Enabled = True
        level2.Enabled = True
        level3.Enabled = False
    End If
End Sub

Sub ordinarymoshi()
    r = MsgBox("您确定要修改游戏模式为经典模式？此操作需要重新开始游戏。", vbOKCancel + vbExclamation + vbSystemModal, "切换模式")
    If r = 1 Then
        ms = 1
        Call tongyong
        ordinary.Enabled = False
        youle.Enabled = True
        prev.Enabled = True
        DeveloperMode.Command10.Enabled = True
        DeveloperMode.Command11.Enabled = True
    End If
End Sub

Sub youlemoshi()
    r = MsgBox("您确定要修改游戏模式为游乐模式？此操作需要重新开始游戏。", vbOKCancel + vbExclamation + vbSystemModal, "切换模式")
    If r = 1 Then
        ms = 2
        Call tongyong
        ordinary.Enabled = True
        youle.Enabled = False
        prev.Enabled = False
        DeveloperMode.Command10.Enabled = False
        DeveloperMode.Command11.Enabled = False
    End If
End Sub

Sub tongyong()
    Erase a
    max = 0: tp = 0: yl = 0: first = 0: js = 0: last = 0
    Call sj
    Call sc
End Sub

Sub baocun_Click()
    Call bc
End Sub
Sub level1_Click()
    Call qingchu
    Call level1a
End Sub

Sub level2_Click()
    Call qingchu
    Call level2a
End Sub

Sub level3_Click()
    Call qingchu
    Call level3a
End Sub

Sub ordinary_Click()
    Call qingchu
    Call ordinarymoshi
End Sub

Sub youle_Click()
    Call qingchu
    Call youlemoshi
End Sub

Sub prev_Click()
    Call huitui
End Sub

Sub nexts_Click()
    Call qianjin
End Sub
Sub pcsong_click()
    Call sxqg
End Sub

Sub MainProgress(keyascii As Integer)
    Erase e
    Erase f
    Erase g
    f3 = False
    f4 = False
    For i = 1 To v * v
        b(i) = a(i)
    Next i
    If keyascii = 97 Or keyascii = 65 Then '向左
        mov1 = 0
        For j = 1 To v
            k = (j - 1) * v + 1
            For i = 1 To v
                m = (j - 1) * v + i
                If a(m) <> 0 Then
                    If m Mod v = 0 Or a(m) < 0 Then
                        n = 0
                    Else
                        For n = m + 1 To j * v
                            If a(n) <> 0 Or n Mod v = 0 Then Exit For
                        Next n
                        i = (n - 1) Mod v
                    End If
                    If a(m) = a(n) Then
                        a(k) = a(m) + 1
                        a(n) = 0
                        Call numchange(2)
                    Else
                        a(k) = a(m)
                        Call numchange(1)
                    End If
                    If m <> k Then a(m) = 0
                    k = k + 1
                End If
            Next i
        Next j
    ElseIf keyascii = 100 Or keyascii = 68 Then '向右
        mov1 = 2
        For j = 1 To v
            k = j * v
            For i = v To 1 Step -1
                m = (j - 1) * v + i
                If a(m) <> 0 Then
                    If m Mod v = 1 Or a(m) < 0 Then
                        n = 0
                    Else
                        For n = m - 1 To (j - 1) * v + 1 Step -1
                            If a(n) <> 0 Or n Mod v = 1 Then Exit For
                        Next n
                        i = n Mod v + 1
                    End If
                    If a(m) = a(n) Then
                        a(k) = a(m) + 1
                        a(n) = 0
                        Call numchange(2)
                    Else
                        a(k) = a(m)
                        Call numchange(1)
                    End If
                    If m <> k Then a(m) = 0
                    k = k - 1
                End If
            Next i
        Next j
    ElseIf keyascii = 119 Or keyascii = 87 Then '向上
        mov1 = 1
        For i = 1 To v
            k = i
            For j = 1 To v
                m = (j - 1) * v + i
                If a(m) <> 0 Then
                    If m > (v - 1) * v Or a(m) < 0 Then
                        n = 0
                    Else
                        For n = m + v To i + (v - 1) * v Step v
                            If a(n) <> 0 Or (n - 1) \ v + 1 = v Then Exit For
                        Next n
                        j = (n - 1) \ v
                    End If
                    If a(m) = a(n) Then
                        a(k) = a(m) + 1
                        a(n) = 0
                        Call numchange(2)
                    Else
                        a(k) = a(m)
                        Call numchange(1)
                    End If
                    If m <> k Then a(m) = 0
                    k = k + v
                End If
            Next j
        Next i
    ElseIf keyascii = 115 Or keyascii = 83 Then '向下
        mov1 = 3
        For i = 1 To v
            k = i + (v - 1) * v
            For j = v To 1 Step -1
                m = (j - 1) * v + i
                If a(m) <> 0 Then
                    If m <= v Or a(m) < 0 Then
                        n = 0
                    Else
                        For n = m - v To i Step -v
                            If a(n) <> 0 Or (n - 1) \ v = 0 Then Exit For
                        Next n
                        j = (n - 1) \ v + 2
                    End If
                    If a(m) = a(n) Then
                        a(k) = a(m) + 1
                        a(n) = 0
                        Call numchange(2)
                    Else
                        a(k) = a(m)
                        Call numchange(1)
                    End If
                    If m <> k Then a(m) = 0
                    k = k - v
                End If
            Next j
        Next i
    End If
    For i = 1 To v * v
        If a(i) <> b(i) Then f3 = True
        If a(i) = 0 Then f4 = True
    Next i
    If f3 = True Or f3 = False And f4 = False Then Call sj
    Call sc
End Sub

Sub numchange(num As Integer)
    ncmax = f(1)
    For nc = 2 To 37
        If f(nc) > ncmax Then ncmax = f(nc)
    Next nc
    If num = 2 Then
        e(ncmax + 1, 2) = n - 1
        If mov1 = 1 Or mov1 = 3 Then
            If v = 4 Then
                g(ncmax + 1, 1) = moshi4.Image1(n - 1).Top
            ElseIf v = 6 Then
                g(ncmax + 1, 1) = moshi6.Image1(n - 1).Top
            End If
        ElseIf mov1 = 0 Or mov1 = 2 Then
            If v = 4 Then
                g(ncmax + 1, 1) = moshi4.Image1(n - 1).Left
            ElseIf v = 6 Then
                g(ncmax + 1, 1) = moshi6.Image1(n - 1).Left
            End If
        End If
    ElseIf num = 1 Then
        e(ncmax + 1, 2) = 0
    End If
    e(ncmax + 1, 1) = m - 1
    f(ncmax + 1) = k - 1
    If mov1 = 1 Or mov1 = 3 Then
        If v = 4 Then
            g(ncmax + 1, 1) = moshi4.Image1(m - 1).Top
        ElseIf v = 6 Then
            g(ncmax + 1, 1) = moshi6.Image1(m - 1).Top
        End If
    ElseIf mov1 = 0 Or mov1 = 2 Then
        If v = 4 Then
            g(ncmax + 1, 1) = moshi4.Image1(m - 1).Left
        ElseIf v = 6 Then
            g(ncmax + 1, 1) = moshi6.Image1(m - 1).Left
        End If
    End If
End Sub

Function newnumber(nd1 As Integer, ms1 As Integer, js1 As Integer) As Integer
    If ms1 = 1 Then '普通模式
        Call ptms
    ElseIf ms1 = 2 Then '游乐模式
        u = (Int(Rnd * 100) + 1) \ ylgl + 1
        If u >= 2 And yl = 0 Then
            tc = -1: yl = yl + 1
            Timer2.Interval = Int(Rnd * 6001) + 1000
            Timer2.Enabled = True
        Else
            Call ptms
        End If
    End If
    newnumber = tc
End Function

Sub ptms()
    If nd = 1 Then
        tc = 1
    ElseIf nd = 2 Then
    tc = Int(Rnd * (js1 \ (100 + js1 \ 10) + 1)) + 1
    ElseIf nd = 3 Then
        tc = Int(Rnd * (js1 \ 100 + 1)) + 1
    End If
End Sub

Private Sub Timer2_Timer()
    Timer2.Enabled = False
    f7 = True
    ylsz = Int(Rnd * 6) + 1
    For i = 1 To v * v
        If a(i) = -1 Then ylwz = i: Exit For
    Next i
    If Timer3.Enabled = True Then Call Delay(0.8)
    If moveclick.Checked = True Then
        DeveloperMode.Command1.Enabled = False
        DeveloperMode.Command2.Enabled = False
        DeveloperMode.Command3.Enabled = False
        If ylwz - v > 0 Then a(ylwz - v) = ylsz: Call mshd(v, ylwz - v): Call zhuangtai: Call Delay(0.8)
        If ylwz Mod v <> 0 Then a(ylwz + 1) = ylsz: Call mshd(v, ylwz + 1): Call zhuangtai:  Call Delay(0.8)
        If ylwz + v <= v * v Then a(ylwz + v) = ylsz: Call mshd(v, ylwz + v):  Call zhuangtai: Call Delay(0.8)
        If ylwz Mod v <> 1 Then a(ylwz - 1) = ylsz: Call mshd(v, ylwz - 1): Call zhuangtai:  Call Delay(0.8)
        Call mshd(v, ylwz, 2): Call zhuangtai: Call Delay(0.8)
        a(ylwz) = 0: yl = 0: Call zhuangtai
        DeveloperMode.Command1.Enabled = True
        DeveloperMode.Command2.Enabled = True
        DeveloperMode.Command3.Enabled = True
    Else
        If ylwz - v > 0 Then a(ylwz - v) = ylsz
        If ylwz + v <= v * v Then a(ylwz + v) = ylsz
        If ylwz Mod v <> 0 Then a(ylwz + 1) = ylsz
        If ylwz Mod v <> 1 Then a(ylwz - 1) = ylsz
        a(ylwz) = 0: yl = 0
        Call qingchu
        Call sc
    End If
    f7 = False
End Sub

Sub mshd(yxdx As Integer, hdzb As Integer, Optional mszt As Integer)
    time = 0
    If mszt <> 2 Then
        If yxdx = 4 Then
            wz1 = moshi4.Image1(hdzb - 1).Width
            dh2 = moshi4.Image1(hdzb - 1).Top + wz1 / 2
            dh3 = moshi4.Image1(hdzb - 1).Left + wz1 / 2
            moshi4.Image1(hdzb - 1).Width = 0
            moshi4.Image1(hdzb - 1).Height = 0
            moshi4.Image1(hdzb - 1) = LoadResPicture(100 + a(hdzb), vbResBitmap)
        ElseIf yxdx = 6 Then
            wz1 = moshi6.Image1(hdzb - 1).Width
            dh2 = moshi6.Image1(hdzb - 1).Top + wz1 / 2
            dh3 = moshi6.Image1(hdzb - 1).Left + wz1 / 2
            moshi6.Image1(hdzb - 1).Width = 0
            moshi6.Image1(hdzb - 1).Height = 0
            moshi6.Image1(hdzb - 1) = LoadResPicture(100 + a(hdzb), vbResBitmap)
        End If
    ElseIf mszt = 2 Then
        If yxdx = 4 Then
            wz1 = moshi4.Image1(hdzb - 1).Width
            dh2 = moshi4.Image1(hdzb - 1).Top + wz1 / 2
            dh3 = moshi4.Image1(hdzb - 1).Left + wz1 / 2
        ElseIf yxdx = 6 Then
            wz1 = moshi6.Image1(hdzb - 1).Width
            dh2 = moshi6.Image1(hdzb - 1).Top + wz1 / 2
            dh3 = moshi6.Image1(hdzb - 1).Left + wz1 / 2
        End If
    End If
    mszt1 = mszt
    hdzb1 = hdzb
    Timer3.Enabled = True
End Sub

Sub sj()
    Randomize
    f1 = False: f2 = False: f6 = False
    k = 0
    Erase c
    For i = 1 To v * v
        If a(i) = 0 Then k = k + 1: c(k) = i
        If a(i) = -1 Then f6 = True
    Next i
    If k <> 0 Then
        t = Int(Rnd() * k) + 1
        tt = c(t)
        a(c(t)) = newnumber(nd, ms, js)
        dh1 = c(t)
        js = js + 1: last = js '当游戏刚开始时会调用一次随机数生成，导致没有移动步数js计数却自增1，因而js的下限为1
        For i = 1 To v * v 'excel屏蔽
            xlSheet.Cells(i + 1, js) = a(i) 'excel屏蔽
        Next i 'excel屏蔽
    ElseIf k = 0 And f6 = False Then
        For i = 1 To v
            For j = 0 To v - 2
                If a(j * v + i) = a((j + 1) * v + i) And a(j * v + i) > 0 Then f1 = True
            Next j
        Next i
        For j = 0 To v - 1
            For i = 1 To v - 1
                If a(j * v + i) = a(j * v + i + 1) And a(j * v + i) > 0 Then f2 = True
            Next i
        Next j
        If f1 = False And f2 = False Then
            For i = 1 To v * v
                If a(i) > max Then max = a(i)
            Next i
            r = MsgBox("模式:" + zfcjc(0, ms) + "   难度:" + zfcjc(nd, 0) + Chr$(10) + "已经没有可以移动或合成的数字了！" + Chr$(10) + "总步数:" + Str(js - 1) + _
            "    最大合成数字:" + Str(2 ^ max), vbRetryCancel + vbExclamation + vbSystemModal, "游戏结束")
            If r = 4 Then '此处进入排行榜程序
                '此处离开排行榜模块
                max = 0
                Erase a
                js = 0
                Call sj
            End If
        End If
    End If
End Sub

Function zfcjc(zfcjc1 As Integer, zfcjc2 As Integer) As String
    If zfcjc1 <> 0 And zfcjc2 = 0 Then
        If zfcjc1 = 1 Then
            zfcjc = "普通"
        ElseIf zfcjc1 = 2 Then
            zfcjc = "困难"
        ElseIf zfcjc1 = 3 Then
            zfcjc = "地狱"
        End If
    ElseIf zfcjc1 = 0 And zfcjc2 <> 0 Then
        If zfcjc2 = 1 Then
            zfcjc = "经典"
        ElseIf zfcjc2 = 2 Then
            zfcjc = "游乐"
        End If
    End If
End Function

Private Sub Timer3_Timer()
    time = time + 15
    If mszt1 <> 2 Or ms = 1 Then
        If v = 4 Then
            If time > wz1 Then
                moshi4.Image1(hdzb1 - 1).Width = wz1
                moshi4.Image1(hdzb1 - 1).Height = wz1
                moshi4.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2
                moshi4.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2
                Timer3.Enabled = False
                Exit Sub
            End If
            moshi4.Image1(hdzb1 - 1).Top = dh2 - time / 2
            moshi4.Image1(hdzb1 - 1).Left = dh3 - time / 2
            moshi4.Image1(hdzb1 - 1).Width = time
            moshi4.Image1(hdzb1 - 1).Height = time
        ElseIf v = 6 Then
            If time > wz1 Then
                moshi6.Image1(hdzb1 - 1).Width = wz1
                moshi6.Image1(hdzb1 - 1).Height = wz1
                moshi6.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2
                moshi6.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2
                Timer3.Enabled = False
                Exit Sub
            End If
            moshi6.Image1(hdzb1 - 1).Top = dh2 - time / 2
            moshi6.Image1(hdzb1 - 1).Left = dh3 - time / 2
            moshi6.Image1(hdzb1 - 1).Width = time
            moshi6.Image1(hdzb1 - 1).Height = time
        End If
    ElseIf mszt1 = 2 And ms = 2 Then
        If v = 4 Then
            If time > wz1 Then
                moshi4.Image1(hdzb1 - 1) = Nothing
                moshi4.Image1(hdzb1 - 1).Width = wz1
                moshi4.Image1(hdzb1 - 1).Height = wz1
                moshi4.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2
                moshi4.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2
                Timer3.Enabled = False
                Exit Sub
            End If
            moshi4.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2 + time / 2
            moshi4.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2 + time / 2
            moshi4.Image1(hdzb1 - 1).Width = wz1 - time
            moshi4.Image1(hdzb1 - 1).Height = wz1 - time
        ElseIf v = 6 Then
            If time > wz1 Then
                moshi6.Image1(hdzb1 - 1) = Nothing
                moshi6.Image1(hdzb1 - 1).Width = wz1
                moshi6.Image1(hdzb1 - 1).Height = wz1
                moshi6.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2
                moshi6.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2
                Timer3.Enabled = False
                Exit Sub
            End If
            moshi6.Image1(hdzb1 - 1).Top = dh2 - wz1 / 2 + time / 2
            moshi6.Image1(hdzb1 - 1).Left = dh3 - wz1 / 2 + time / 2
            moshi6.Image1(hdzb1 - 1).Width = wz1 - time
            moshi6.Image1(hdzb1 - 1).Height = wz1 - time
        End If
    End If
End Sub

Sub sc()
    f7 = True
    If f3 = True And moveclick.Checked = True Then
        For i = 1 To 100
            For count1 = 1 To ncmax + 1
                If e(count1, 2) <> 0 Then
                    If mov1 = 1 Then
                        If v = 4 Then
                            If moshi4.Image1(e(count1, 2)).Top <> moshi4.Image1(f(count1)).Top Then
                                moshi4.Image1(e(count1, 2)).Top = moshi4.Image1(e(count1, 2)).Top - Abs(moshi4.Image1(f(count1)).Top - g(count1, 2)) / 100
                            End If
                        ElseIf v = 6 Then
                            If moshi6.Image1(e(count1, 2)).Top <> moshi6.Image1(f(count1)).Top Then
                                moshi6.Image1(e(count1, 2)).Top = moshi6.Image1(e(count1, 2)).Top - Abs(moshi6.Image1(f(count1)).Top - g(count1, 2)) / 100
                            End If
                        End If
                    ElseIf mov1 = 3 Then
                        If v = 4 Then
                            If moshi4.Image1(e(count1, 2)).Top <> moshi4.Image1(f(count1)).Top Then
                                moshi4.Image1(e(count1, 2)).Top = moshi4.Image1(e(count1, 2)).Top + Abs(moshi4.Image1(f(count1)).Top - g(count1, 2)) / 100
                            End If
                        ElseIf v = 6 Then
                            If moshi6.Image1(e(count1, 2)).Top <> moshi6.Image1(f(count1)).Top Then
                                moshi6.Image1(e(count1, 2)).Top = moshi6.Image1(e(count1, 2)).Top + Abs(moshi6.Image1(f(count1)).Top - g(count1, 2)) / 100
                            End If
                        End If
                    ElseIf mov1 = 0 Then
                        If v = 4 Then
                            If moshi4.Image1(e(count1, 2)).Left <> moshi4.Image1(f(count1)).Left Then
                                moshi4.Image1(e(count1, 2)).Left = moshi4.Image1(e(count1, 2)).Left - Abs(moshi4.Image1(f(count1)).Left - g(count1, 2)) / 100
                            End If
                        ElseIf v = 6 Then
                            If moshi6.Image1(e(count1, 2)).Left <> moshi6.Image1(f(count1)).Left Then
                                moshi6.Image1(e(count1, 2)).Left = moshi6.Image1(e(count1, 2)).Left - Abs(moshi6.Image1(f(count1)).Left - g(count1, 2)) / 100
                            End If
                        End If
                    ElseIf mov1 = 2 Then
                        If v = 4 Then
                            If moshi4.Image1(e(count1, 2)).Left <> moshi4.Image1(f(count1)).Left Then
                                moshi4.Image1(e(count1, 2)).Left = moshi4.Image1(e(count1, 2)).Left + Abs(moshi4.Image1(f(count1)).Left - g(count1, 2)) / 100
                            End If
                        ElseIf v = 6 Then
                            If moshi6.Image1(e(count1, 2)).Left <> moshi6.Image1(f(count1)).Left Then
                                moshi6.Image1(e(count1, 2)).Left = moshi6.Image1(e(count1, 2)).Left + Abs(moshi6.Image1(f(count1)).Left - g(count1, 2)) / 100
                            End If
                        End If
                    End If
                End If
                If mov1 = 1 Then
                    If v = 4 Then
                        If moshi4.Image1(e(count1, 1)).Top <> moshi4.Image1(f(count1)).Top Then
                            moshi4.Image1(e(count1, 1)).Top = moshi4.Image1(e(count1, 1)).Top - Abs(moshi4.Image1(f(count1)).Top - g(count1, 1)) / 100
                        End If
                    ElseIf v = 6 Then
                        If moshi6.Image1(e(count1, 1)).Top <> moshi6.Image1(f(count1)).Top Then
                            moshi6.Image1(e(count1, 1)).Top = moshi6.Image1(e(count1, 1)).Top - Abs(moshi6.Image1(f(count1)).Top - g(count1, 1)) / 100
                        End If
                    End If
                ElseIf mov1 = 3 Then
                    If v = 4 Then
                        If moshi4.Image1(e(count1, 1)).Top <> moshi4.Image1(f(count1)).Top Then
                            moshi4.Image1(e(count1, 1)).Top = moshi4.Image1(e(count1, 1)).Top + Abs(moshi4.Image1(f(count1)).Top - g(count1, 1)) / 100
                        End If
                    ElseIf v = 6 Then
                        If moshi6.Image1(e(count1, 1)).Top <> moshi6.Image1(f(count1)).Top Then
                            moshi6.Image1(e(count1, 1)).Top = moshi6.Image1(e(count1, 1)).Top + Abs(moshi6.Image1(f(count1)).Top - g(count1, 1)) / 100
                        End If
                    End If
                ElseIf mov1 = 0 Then
                    If v = 4 Then
                        If moshi4.Image1(e(count1, 1)).Left <> moshi4.Image1(f(count1)).Left Then
                            moshi4.Image1(e(count1, 1)).Left = moshi4.Image1(e(count1, 1)).Left - Abs(moshi4.Image1(f(count1)).Left - g(count1, 1)) / 100
                        End If
                    ElseIf v = 6 Then
                        If moshi6.Image1(e(count1, 1)).Left <> moshi6.Image1(f(count1)).Left Then
                            moshi6.Image1(e(count1, 1)).Left = moshi6.Image1(e(count1, 1)).Left - Abs(moshi6.Image1(f(count1)).Left - g(count1, 1)) / 100
                        End If
                    End If
                ElseIf mov1 = 2 Then
                    If v = 4 Then
                        If moshi4.Image1(e(count1, 1)).Left <> moshi4.Image1(f(count1)).Left Then
                            moshi4.Image1(e(count1, 1)).Left = moshi4.Image1(e(count1, 1)).Left + Abs(moshi4.Image1(f(count1)).Left - g(count1, 1)) / 100
                        End If
                    ElseIf v = 6 Then
                        If moshi6.Image1(e(count1, 1)).Left <> moshi6.Image1(f(count1)).Left Then
                            moshi6.Image1(e(count1, 1)).Left = moshi6.Image1(e(count1, 1)).Left + Abs(moshi6.Image1(f(count1)).Left - g(count1, 1)) / 100
                        End If
                    End If
                End If
            Next count1
        Next i
        If v = 4 Then
            For i = 1 To ncmax + 1
                If e(count1, 2) <> 0 Then moshi4.Image1(e(count1, 2)) = Nothing
            Next i
            For i = 0 To 15
                moshi4.Image1(i).Visible = False
            Next i
            For i = 0 To 3
                For j = 0 To 3
                    moshi4.Image1(i * 4 + j).Top = 152 * i
                    moshi4.Image1(i * 4 + j).Left = 152 * j
                Next j
            Next i
            For i = 0 To 15
                moshi4.Image1(i).Visible = True
            Next i
        ElseIf v = 6 Then
            For i = 1 To ncmax + 1
                If e(count1, 2) <> 0 Then moshi6.Image1(e(count1, 2)) = Nothing
            Next i
            For i = 0 To 35
                moshi6.Image1(i).Visible = False
            Next i
            For i = 0 To 5
                For j = 0 To 5
                    moshi6.Image1(i * 6 + j).Top = 104 * i
                    moshi6.Image1(i * 6 + j).Left = 104 * j
                Next j
            Next i
            For i = 0 To 35
                moshi6.Image1(i).Visible = True
            Next i
        End If
    End If
    For i = 0 To v * v - 1
        If a(i + 1) <> b(i + 1) Or i + 1 = tt Then
            If v = 4 Then
                If a(i + 1) = 0 Then
                    moshi4.Image1(i).Picture = Nothing
                Else
                    If i + 1 <> dh1 Or moveclick.Checked = False Then
                        moshi4.Image1(i) = LoadResPicture(100 + a(i + 1), vbResBitmap)
                    Else
                        Call mshd(v, dh1)
                    End If
                End If
            ElseIf v = 6 Then
                If a(i + 1) = 0 Then
                    moshi6.Image1(i).Picture = Nothing
                Else
                    If i + 1 <> dh1 Or moveclick.Checked = False Then
                        moshi6.Image1(i) = LoadResPicture(100 + a(i + 1), vbResBitmap)
                    Else
                        Call mshd(v, dh1)
                    End If
                End If
            End If
        End If
        tt = -2
    Next i
    Call zhuangtai
    f7 = False
End Sub

Sub zhuangtai()
    If v <> 0 Then
        DeveloperMode.List1.Clear
        If v <> 5 Then
            For i = 1 To v * v
                s = s + " a(" + zfc(i) + ")=" + zfc(a(i))
                If i Mod v = 0 Then
                    DeveloperMode.List1.AddItem s
                    s = ""
                End If
            Next i
            Call jiance
        End If
    End If
End Sub

Sub jiance() '全屏与非全屏不同的过程
    DeveloperMode.List2.Clear
    If v <> 5 Then
        If first = 0 Then
            DeveloperMode.List2.AddItem "起始(first)=" + Str(first)
        Else
            DeveloperMode.List2.AddItem "起始(first)=" + Str(first - 1)
        End If
        DeveloperMode.List2.AddItem "步数（js）=" + Str(js - 1)
        DeveloperMode.List2.AddItem "末尾（last）=" + Str(last - 1)
        DeveloperMode.List2.AddItem "难度（nd）=" + Str(nd)
        DeveloperMode.List2.AddItem "模式（ms）=" + Str(ms)
        DeveloperMode.List2.AddItem "音乐（yy）=" + Str(yy)
    End If
End Sub

Public Function zfc(sz As Integer) As String
    If Len(CStr(sz)) = 1 Then
        zfc = " " + CStr(sz)
    Else
        zfc = CStr(sz)
    End If
End Function

Sub loadgame_Click() '全屏与非全屏不同的过程
    xrsc = ""
    With CommonDialog1
        .DialogTitle = "读取存档"
        .Filter = "文本文件(*.txt)|*.txt|所有文件(*.*)|*.*"
        .FilterIndex = 0
        .InitDir = App.Path & "\saves"
        .ShowOpen
        Dim tmpLoadStr As String
        If .FileName = "" Then Exit Sub
        Open .FileName For Input As #1
            Do While Not EOF(1)
                Line Input #1, tmpLoadStr
                xrsc = xrsc & tmpLoadStr & vbCrLf
            Loop
        Close #1
    End With
    u = 0: t = 0
    For i = 1 To Len(xrsc)
        If Not ((0 <= Val(Mid(xrsc, i, 1)) And Val(Mid(xrsc, i, 1))) <= 9 Or Mid(xrsc, i, 1) = "/") Then
            MsgBox "存档数据错误 代码:0X00-Return(" + Trim(Str(i)) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败"
            Exit Sub
        End If
    Next i
    For i = 1 To Len(xrsc)
        If Mid(xrsc, i, 1) = "/" Then
            u = u + 1
            d(u) = Val(Mid(xrsc, t + 1, i - t - 1))
            t = i
        End If
    Next i
    If u <> 40 Then
        MsgBox "存档数据错误 代码:0X01-Return(" + Trim(Str(u)) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败"
        Exit Sub
    End If
    If d(1) <> 4 And d(1) <> 6 Then MsgBox "存档数据错误 代码:0X02-Return(" + Trim(Str(d(1))) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败": Exit Sub
    If d(2) < 1 Or d(2) > 2 Then MsgBox "存档数据错误 代码:0X03-Return(" + Trim(Str(d(2))) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败": Exit Sub
    If d(3) < 1 Or d(3) > 3 Then MsgBox "存档数据错误 代码:0X04-Return(" + Trim(Str(d(3))) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败": Exit Sub
    If d(4) = 0 Then MsgBox "存档数据错误 代码:0X05，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败": Exit Sub
    For i = 6 To 41
        If d(i) > 17 Or d(i) < -1 Then
            MsgBox "存档数据错误 代码:0X06-Return(" + Trim(Str(i)) + ")，读取失败", 0 + vbExclamation + vbSystemModal, "读取失败"
            Exit Sub
        End If
    Next i
    v = d(1)
    ms = d(2)
    nd = d(3)
    js = d(4)
    first = js
    last = js
    For i = 1 To 36
        a(i) = d(i + 4)
    Next i
    MsgBox "读取成功", vbSystemModal
    Call qingchu
    If v = 4 Then
        Unload moshi6
        moshi4.Show 0
        moshi4.Move Screen.Width - moshi4.Width, 0
    ElseIf v = 6 Then
        Unload moshi4
        moshi6.Show 0
        moshi6.Move Screen.Width - moshi6.Width, 0
    End If
    If nd = 1 Then
        level1.Enabled = False
        level2.Enabled = True
        level3.Enabled = True
    ElseIf nd = 2 Then
        level1.Enabled = True
        level2.Enabled = False
        level3.Enabled = True
    ElseIf nd = 3 Then
        level1.Enabled = True
        level2.Enabled = True
        level3.Enabled = False
    End If
    If ms = 1 Then
        ordinary.Enabled = False
        youle.Enabled = True
        prev.Enabled = True
        nexts.Enabled = True
        DeveloperMode.Command10.Enabled = True
        DeveloperMode.Command11.Enabled = True
    ElseIf ms = 2 Then
        ordinary.Enabled = True
        youle.Enabled = False
        prev.Enabled = False
        nexts.Enabled = False
        DeveloperMode.Command10.Enabled = False
        DeveloperMode.Command11.Enabled = False
    End If
    Call sc
End Sub

Sub bc()
    For i = 1 To 36
        If a(i) = -1 Then
            MsgBox "当前状态存在特殊元素，保存失败", vbSystemModal
            Exit Sub
        End If
    Next i
    xrsc = "": returntime = ""
    xrsc = xrsc + CStr(v) + "/" + CStr(ms) + "/" + CStr(nd) + "/" + CStr(js) + "/"
    For i = 1 To 36
        xrsc = xrsc + CStr(a(i)) + "/"
    Next i
    With CommonDialog1
        returntime = Format(Now(), "YYYYmmDDHHMMss")
        Open App.Path & "\saves\" & returntime & ".txt" For Output As #1
            Print #1, xrsc
        Close #1
    End With
    MsgBox "已成功将游戏进度保存至 " & App.Path & "\saves\" & returntime & ".txt", vbSystemModal
End Sub

Sub moveclick_Click()
    f5 = Not f5
    moveclick.Checked = Not moveclick.Checked
End Sub

Private Sub Timer1_Timer()
    If Me.WindowsMediaPlayer1.playState = 1 Then
        Call music
        Call jiance
        WindowsMediaPlayer1.URL = App.Path & musiclist(yy)
    End If
End Sub

Private Sub update_Click()
     UpdateForm.Show 1
End Sub

Sub tmda()
    Dim rtn As Long
    rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hwnd, GWL_EXSTYLE, rtn
    SetLayeredWindowAttributes hwnd, 0, MainForm.tmd1, LWA_ALPHA
End Sub

Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu help
    End If
End Sub

Sub Form_KeyDown(KeyCode As Integer, Shift As Integer) '全屏与非全屏不同的过程
    If KeyCode = 27 Then
        myexit = MsgBox("您确定要退出游戏？", vbExclamation + vbYesNo + vbDefaultButton2 + vbSystemModal, "退出之二")
        If myexit = vbYes Then
            Set xlSheet = Nothing 'excel屏蔽
            xlBook.Close (True) 'excel屏蔽
            Set xlBook = Nothing 'excel屏蔽
            xlApp.DisplayAlerts = False 'excel屏蔽
            xlApp.Quit 'excel屏蔽
            Set xlApp = Nothing 'excel屏蔽
            propath ("EXCEL.EXE") 'excel屏蔽
            Unload MainForm
            Unload moshi4
            Unload moshi6
            Unload duorenmoshi
            Unload DeveloperMode
            Unload Loading
            Unload UpdateForm
            Unload LoginScreen
            Unload about
        End If
    End If
End Sub
