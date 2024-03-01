VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form about 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "游戏说明"
   ClientHeight    =   8040
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   11790
   Icon            =   "Form8.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "关闭，下次仍然提示"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   7560
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "关闭，下次不再提示"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   5895
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   7575
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form8.frx":038A
      Top             =   0
      Width           =   11775
   End
End
Attribute VB_Name = "about"
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
    aboutopen = Trim(Str(0))
    With CommonDialog1
        Open App.Path & "\flags\aboutflag.txt" For Output As #1
        Print #1, aboutopen
        Close #1
    End With
    Unload about
End Sub

Private Sub Command2_Click()
    aboutopen = Trim(Str(1))
    With CommonDialog1
        Open App.Path & "\flags\aboutflag.txt" For Output As #1
        Print #1, aboutopen
        Close #1
    End With
    Unload about
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Text1.Text = "2048基本规则：控制所有方块向同一方向移动，在该方向的逆向上先配对的两个相邻或隔空的相同数字方块会合并成为它们的和，若没有配对则保持不变，全部合并完成后沿该方向沉底，如果所有格子塞满并且无方块可以合并则判定游戏结束" _
+ vbCrLf + vbCrLf + "在游戏窗口右键可以呼出游戏菜单，包含上下步移动、保存读取存档（保存至游戏文件夹\saves\以年月日时分秒命名的txt文本文件）、开启/关闭滑动开关、调整游戏模式（经典、游乐）、调整游戏难度（普通、困难、地狱）以及顺序切歌（共29首）" + vbCrLf + vbCrLf + "在主窗体上右键可以呼出信息菜单，包含游戏说明（本窗体）以及2048更新日志" _
+ vbCrLf + vbCrLf + "只要更改了游戏难度或者游戏模式都将重新开始，在困难难度及以上的难度下生成的数字会与步数相关联，使游戏增添了一份玄学的气息" + vbCrLf + Chr$(10) + "【新版游乐模式】在游乐模式下你行动的每一步都有概率（默认为5%）随机在某一个格子出现游乐，并且会在随机的时间内向四周生成相同的随机大小数字给玩家带来希望或者绝望" + vbCrLf + "游乐，永远滴神！（0~0）" _
+ vbCrLf + vbCrLf + "四六游戏操作:" + vbCrLf + "双击图标4或6进入不同大小游戏窗体游戏，满格且无法移动后游戏结束" _
+ vbCrLf + vbCrLf + "多人游戏操作:" + vbCrLf + "双击无穷标志进入多人对战，在文本框中输入数字后回车或按下方按钮以开始（数字不是时长），等到上方进度条变成纯蓝色且最长时开放双方按钮（不区分大小写的WASD和上下左右方向键）进行对战，待进度条变红消失在中央圆或双方都无法移动数字时游戏结束，以双方合成的最大值作为比较胜负的标准，在左上角会打印出本次对战的始末信息" _
+ vbCrLf + vbCrLf + "退出游戏操作:" + vbCrLf + "在游戏窗体上按ESC会弹出是否保存存档，是/否后会关闭游戏窗体，在主窗体上按ESC会弹出是否退出游戏，确认后正式退出" _
+ vbCrLf + vbCrLf + "操作方法:" + vbCrLf + "1-于游戏窗体右键激活滑动开关，直接使用鼠标按住-滑动-松开或者使用触摸屏点划" + vbCrLf + "2-激活要玩耍的窗体，按动键盘（不区分大小写，但不能是中文输入法）" + vbCrLf + "  W——向上" + vbCrLf + "  A——向左" + vbCrLf + "  S——向下" + vbCrLf + "  D——向右" _
+ vbCrLf + vbCrLf + "当前版本:2048V6.0.0正式版" + vbCrLf + "游戏制作:裂空之云&裂云之空" + vbCrLf + "本软件全部权利归裂空之云&裂云之空所有" + vbCrLf + "未获得作者许可，禁止对本软件进行任何形式的发布、修改、反编译、破解等行为" + vbCrLf + "本软件仅供日常娱乐，禁止将本软件用于任何形式的商业行为"
End Sub
