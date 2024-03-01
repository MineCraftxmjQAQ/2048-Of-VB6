VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Loading 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17055
   FillColor       =   &H80000001&
   Icon            =   "Form6.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   602
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1137
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   120
      _ExtentX        =   212
      _ExtentY        =   529
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      Appearance      =   0
      TextRTF         =   $"Form6.frx":038A
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   11295
      Left            =   -1920
      Picture         =   "Form6.frx":0427
      ScaleHeight     =   749
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1413
      TabIndex        =   0
      Top             =   -1080
      Visible         =   0   'False
      Width           =   21255
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   2292
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   3492
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
         _cx             =   6165
         _cy             =   4048
      End
   End
End
Attribute VB_Name = "Loading"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lTime As Byte
Dim x As Double
Dim color1, color2, color3 As Byte
Dim sinx As Double
Const pi = 3.1415926

Private Sub Form_Load() '全屏与非全屏不同的过程
    If App.PrevInstance = True Then
        MsgBox "你已经打开了一个游戏，不能打开第二个。", 0 + vbExclamation + vbSystemModal, "禁止重复打开"
        End
    Else
        ScrX = Screen.Width / Screen.TwipsPerPixelX '获取屏幕宽度
        ScrY = Screen.Height / Screen.TwipsPerPixelY '获取屏幕高度
        Loading.Show
        RichTextBox1.Visible = True
        RichTextBox1.Width = 0
        Call Delay(0.4)
        WindowsMediaPlayer1.URL = App.Path & "\music\Never Be Alone (Michael Jan Remix).mp3"
        RichTextBox1.Width = ScrX * 0.2
        RichTextBox1.Text = "播放音乐"
        Call Delay(0.4)
        Set xlApp = CreateObject("Excel.Application") 'excel屏蔽
        RichTextBox1.Width = ScrX * 0.4
        RichTextBox1.Text = "正在连接Excel"
        Call Delay(0.4)
        Set xlBook = xlApp.Workbooks.Open(App.Path & "\xls\Historical records.xlsx") 'excel屏蔽
        RichTextBox1.Width = ScrX * 0.6
        RichTextBox1.Text = "正在连接历史记录表"
        Call Delay(0.4)
        Set xlSheet = xlBook.Worksheets("sheet1") 'excel屏蔽
        RichTextBox1.Width = ScrX * 0.8
        RichTextBox1.Text = "正在连接历史记录"
        Call Delay(0.4)
        Load MainForm
        MainForm.WindowsMediaPlayer1.Controls.stop
        RichTextBox1.Width = ScrX
        RichTextBox1.Text = "正在加载主交互界面"
        Call Delay(0.4)
        lTime = 0
        x = 0
        Timer1.Enabled = True
    End If
End Sub

Sub ShowTransparency(cSrc As PictureBox, cDest As Form, ByVal nLevel As Byte)
    Dim LrProps As rBlendProps
    Dim LnBlendPtr As Long
    cDest.Cls
    LrProps.tBlendAmount = nLevel
    CopyMemory LnBlendPtr, LrProps, 4
    With cSrc
        AlphaBlend cDest.hDC, (Loading.ScaleWidth - Picture1.Width) / 2, (Loading.ScaleHeight - Picture1.Height) / 2, .ScaleWidth, .ScaleHeight, .hDC, 0, 0, .ScaleWidth, .ScaleHeight, LnBlendPtr
    End With
    cDest.Refresh
End Sub

Sub RichTextBox1_Keydown(KeyCode As Integer, Shift As Integer)
    If lTime = 255 Then
        Me.WindowsMediaPlayer1.Controls.stop
        Unload Loading
        MainForm.WindowsMediaPlayer1.Controls.play
        MainForm.Show
        Open App.Path & "\flags\aboutflag.txt" For Input As #1 '下面打开帮助文本
        Line Input #1, aboutopen
        Close #1
        If Not (Len(aboutopen) = 1 And (Val(aboutopen) = 1 Or Val(aboutopen) = 0)) Then
            aboutopen = Trim(Str(1))
            With CommonDialog1
                Open App.Path & "\flags\aboutflag.txt" For Output As #1
                Print #1, aboutopen
                Close #1
            End With
        End If
        If Val(aboutopen) = 1 Then about.Show 1 '此处结束
    End If
End Sub

Sub Form_Click()
    Dim af As String
    If lTime = 255 Then
        Me.WindowsMediaPlayer1.Controls.stop
        Unload Loading
        MainForm.WindowsMediaPlayer1.Controls.play
        MainForm.Show
        With CommonDialog1
            Open App.Path & "\flags\aboutflag.txt" For Input As #1
            If Not EOF(1) Then Line Input #1, af
                If af = "1" Then
                    Load about
                    about.Show
                End If
            Close #1
        End With
    End If
End Sub

Private Sub RichTextBox1_Change()
    If lTime <= 85 And lTime > 0 Then
        RichTextBox1.SelStart = 4
        RichTextBox1.SelLength = 4
        RichTextBox1.SelColor = RGB(255 - lTime, lTime, 2 * lTime)
        RichTextBox1.SelStart = 0
        RichTextBox1.SelLength = 0
    ElseIf lTime <= 170 Then
        RichTextBox1.SelStart = 4
        RichTextBox1.SelLength = 4
        RichTextBox1.SelColor = RGB(170 - lTime / 4, 85 + lTime \ 4, 255 - lTime)
        RichTextBox1.SelStart = 0
        RichTextBox1.SelLength = 0
    ElseIf lTime < 225 Then
        RichTextBox1.SelStart = 4
        RichTextBox1.SelLength = 4
        RichTextBox1.SelColor = RGB(lTime, lTime \ 2, 255 - lTime)
        RichTextBox1.SelStart = 0
        RichTextBox1.SelLength = 0
    ElseIf lTime = 255 Then
        RichTextBox1.SelStart = 0
        RichTextBox1.SelLength = Len(RichTextBox1.Text)
        RichTextBox1.SelColor = RGB(color1, color2, color3)
        RichTextBox1.SelStart = 0
        RichTextBox1.SelLength = 0
    End If
End Sub
Private Sub Timer1_Timer() '全屏与非全屏不同的过程
    If lTime < 254 Then
        lTime = lTime + 2
        ShowTransparency Picture1, Loading, lTime
        RichTextBox1.Text = "正在加载" + Str(Int(lTime / 2.55)) + "%"
        RichTextBox1.SelColor = vbRed
        RichTextBox1.Width = ScrX * lTime / 255
    ElseIf lTime = 254 Then
        lTime = lTime + 1
    End If
    If Int(lTime / 2.55) = 100 Then
        Picture1.Picture = LoadResPicture(118, 0)
        RichTextBox1.Width = ScrX
        RichTextBox1.Text = "加载完成，点击任意位置进入游戏。"
        sinx = 255 * (0.2 * Cos(0.04 * x) + 0.8)
        color1 = Int(255 * (0.45 * Cos(0.02 * x) + 0.55))
        color2 = Int(255 * (1 / pi * Abs(ArcSin(Sin(0.01 * pi * x))) + 0.4))
        color3 = Int(255 * (0.45 * Sin(0.02 * x) + 0.55))
        x = x + 1
        ShowTransparency Picture1, Loading, sinx
    End If
End Sub
