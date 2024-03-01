VERSION 5.00
Begin VB.Form duorenmoshi 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "双人对战"
   ClientHeight    =   9360
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   14880
   Icon            =   "Form9.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   624
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   992
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   1332
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000B&
      Caption         =   "输入时长后点击此按钮开始"
      Height          =   492
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   1332
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   252
      Left            =   0
      Top             =   1920
      Width           =   14892
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   7212
      Left            =   7320
      Top             =   2160
      Width           =   252
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1200
      Left            =   6840
      Shape           =   3  'Circle
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   15
      Left            =   13080
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   14
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   13
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   12
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   11
      Left            =   13080
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   10
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   9
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   8
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   7
      Left            =   13080
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   6
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   5
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   4
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   3
      Left            =   13080
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   2
      Left            =   11280
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   1
      Left            =   9480
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image2 
      Height          =   1800
      Index           =   0
      Left            =   7680
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   15
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   14
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   13
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   12
      Left            =   0
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   11
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   10
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   9
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   8
      Left            =   0
      Stretch         =   -1  'True
      Top             =   5760
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   7
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   6
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   5
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   4
      Left            =   0
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   3
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   2
      Left            =   3600
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   1
      Left            =   1800
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Image Image1 
      Height          =   1800
      Index           =   0
      Left            =   0
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1800
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   372
      Left            =   7200
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Visible         =   0   'False
      Width           =   492
   End
End
Attribute VB_Name = "duorenmoshi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST& = -1
Private Const SWP_NOSIZE& = &H1
Private Const SWP_NOMOVE& = &H2
Dim e(0 To 16) As Integer
Dim f(0 To 16) As Integer
Dim s1(1 To 16) As Integer
Dim i1 As Integer
Dim j1 As Integer
Dim m1 As Integer
Dim n1 As Integer
Dim k1 As Integer
Dim t1 As Integer
Dim dh1 As Integer
Dim max1 As Integer
Dim fa As Boolean
Dim fb As Boolean
Dim fx1 As Boolean
Dim fy1 As Boolean
Dim s3 As String

Dim g(0 To 16) As Integer
Dim h(0 To 16) As Integer
Dim s2(1 To 16) As Integer
Dim i2 As Integer
Dim j2 As Integer
Dim m2 As Integer
Dim n2 As Integer
Dim k2 As Integer
Dim t2 As Integer
Dim t3 As Integer
Dim t4 As Integer
Dim dh2 As Integer
Dim max2 As Integer
Dim fc As Boolean
Dim fd As Boolean
Dim fx2 As Boolean
Dim fy2 As Boolean
Dim s4 As String

Dim ur As Integer
Dim shijian As Integer
Dim jieshu As Integer

Dim f5 As Boolean



Sub MainProgress1(keyascii As Integer)
    fa = False
    fb = False
    For i1 = 1 To 16
        f(i1) = e(i1)
    Next i1
    If keyascii = 97 Or keyascii = 65 Then '向左
        For j1 = 1 To 4
           k1 = (j1 - 1) * 4 + 1
            For i1 = 1 To 4
                m1 = (j1 - 1) * 4 + i1
                If e(m1) <> 0 Then
                    If m1 Mod 4 = 0 Then
                        n1 = 0
                    Else
                        For n1 = m1 + 1 To j1 * 4
                            If e(n1) <> 0 Or n1 Mod 4 = 0 Then Exit For
                        Next n1
                        i1 = (n1 - 1) Mod 4
                    End If
                    If e(m1) = e(n1) Then
                        e(k1) = e(m1) + 1
                        e(n1) = 0
                    Else
                        e(k1) = e(m1)
                    End If
                    If m1 <> k1 Then e(m1) = 0
                   k1 = k1 + 1
                End If
            Next i1
        Next j1
    ElseIf keyascii = 100 Or keyascii = 68 Then '向右
        For j1 = 1 To 4
           k1 = j1 * 4
            For i1 = 4 To 1 Step -1
                m1 = (j1 - 1) * 4 + i1
                If e(m1) <> 0 Then
                    If m1 Mod 4 = 1 Then
                        n1 = 0
                    Else
                        For n1 = m1 - 1 To (j1 - 1) * 4 + 1 Step -1
                            If e(n1) <> 0 Or n1 Mod 4 = 1 Then Exit For
                        Next n1
                        i1 = n1 Mod 4 + 1
                    End If
                    If e(m1) = e(n1) And n1 > (j1 - 1) * 4 Then
                        e(k1) = e(m1) + 1
                        e(n1) = 0
                    Else
                        e(k1) = e(m1)
                    End If
                    If m1 <> k1 Then e(m1) = 0
                   k1 = k1 - 1
                End If
            Next i1
        Next j1
    ElseIf keyascii = 119 Or keyascii = 87 Then '向上
        For i1 = 1 To 4
           k1 = i1
            For j1 = 1 To 4
                m1 = (j1 - 1) * 4 + i1
                If e(m1) <> 0 Then
                    If m1 > 12 Then
                        n1 = 0
                    Else
                        For n1 = m1 + 4 To i1 + 12 Step 4
                            If e(n1) <> 0 Or (n1 - 1) \ 4 + 1 = 4 Then Exit For
                        Next n1
                        j1 = (n1 - 1) \ 4
                    End If
                    If e(m1) = e(n1) Then
                        e(k1) = e(m1) + 1
                        e(n1) = 0
                    Else
                        e(k1) = e(m1)
                    End If
                    If m1 <> k1 Then e(m1) = 0
                   k1 = k1 + 4
                End If
            Next j1
        Next i1
    ElseIf keyascii = 115 Or keyascii = 83 Then '向下
        For i1 = 1 To 4
           k1 = i1 + 12
            For j1 = 4 To 1 Step -1
                m1 = (j1 - 1) * 4 + i1
                If e(m1) <> 0 Then
                    If m1 <= 4 Then
                        n1 = 0
                    Else
                        For n1 = m1 - 4 To i1 Step -4
                            If e(n1) <> 0 Or (n1 - 1) \ 4 = 0 Then Exit For
                        Next n1
                        j1 = (n1 - 1) \ 4 + 2
                    End If
                    If e(m1) = e(n1) Then
                        e(k1) = e(m1) + 1
                        e(n1) = 0
                    Else
                        e(k1) = e(m1)
                    End If
                    If m1 <> k1 Then e(m1) = 0
                   k1 = k1 - 4
                End If
            Next j1
        Next i1
    End If
    For i1 = 1 To 16
        If e(i1) <> f(i1) Then fa = True
        If e(i1) = 0 Then fb = True
    Next i1
    If fa = True Or fa = False And fb = False Then Call sj1
    Call sc1
End Sub

Sub MainProgress2(KeyCode As Integer)
    fc = False
    fd = False
    For i2 = 1 To 16
        h(i2) = g(i2)
    Next i2
    If KeyCode = 37 Then '向左
        For j2 = 1 To 4
            k2 = (j2 - 1) * 4 + 1
            For i2 = 1 To 4
                m2 = (j2 - 1) * 4 + i2
                If g(m2) <> 0 Then
                    If m2 Mod 4 = 0 Then
                        n2 = 0
                    Else
                        For n2 = m2 + 1 To j2 * 4
                            If g(n2) <> 0 Or n2 Mod 4 = 0 Then Exit For
                        Next n2
                        i2 = (n2 - 1) Mod 4
                    End If
                    If g(m2) = g(n2) Then
                        g(k2) = g(m2) + 1
                        g(n2) = 0
                    Else
                        g(k2) = g(m2)
                    End If
                    If m2 <> k2 Then g(m2) = 0
                    k2 = k2 + 1
                End If
            Next i2
        Next j2
    ElseIf KeyCode = 39 Then '向右
        For j2 = 1 To 4
            k2 = j2 * 4
            For i2 = 4 To 1 Step -1
                m2 = (j2 - 1) * 4 + i2
                If g(m2) <> 0 Then
                    If m2 Mod 4 = 1 Then
                        n2 = 0
                    Else
                        For n2 = m2 - 1 To (j2 - 1) * 4 + 1 Step -1
                            If g(n2) <> 0 Or n2 Mod 4 = 1 Then Exit For
                        Next n2
                        i2 = n2 Mod 4 + 1
                    End If
                    If g(m2) = g(n2) And n2 > (j2 - 1) * 4 Then
                        g(k2) = g(m2) + 1
                        g(n2) = 0
                    Else
                        g(k2) = g(m2)
                    End If
                    If m2 <> k2 Then g(m2) = 0
                    k2 = k2 - 1
                End If
            Next i2
        Next j2
    ElseIf KeyCode = 38 Then '向上
        For i2 = 1 To 4
            k2 = i2
            For j2 = 1 To 4
                m2 = (j2 - 1) * 4 + i2
                If g(m2) <> 0 Then
                    If m2 > (4 - 1) * 4 Then
                        n2 = 0
                    Else
                        For n2 = m2 + 4 To i2 + 12 Step 4
                            If g(n2) <> 0 Or (n2 - 1) \ 4 + 1 = 4 Then Exit For
                        Next n2
                        j2 = (n2 - 1) \ 4
                    End If
                    If g(m2) = g(n2) Then
                        g(k2) = g(m2) + 1
                        g(n2) = 0
                    Else
                        g(k2) = g(m2)
                    End If
                    If m2 <> k2 Then g(m2) = 0
                    k2 = k2 + 4
                End If
            Next j2
        Next i2
    ElseIf KeyCode = 40 Then '向下
        For i2 = 1 To 4
            k2 = i2 + 12
            For j2 = 4 To 1 Step -1
                m2 = (j2 - 1) * 4 + i2
                If g(m2) <> 0 Then
                    If m2 <= 4 Then
                        n2 = 0
                    Else
                        For n2 = m2 - 4 To i2 Step -4
                            If g(n2) <> 0 Or (n2 - 1) \ 4 = 0 Then Exit For
                        Next n2
                        j2 = (n2 - 1) \ 4 + 2
                    End If
                    If g(m2) = g(n2) Then
                        g(k2) = g(m2) + 1
                        g(n2) = 0
                    Else
                        g(k2) = g(m2)
                    End If
                    If m2 <> k2 Then g(m2) = 0
                    k2 = k2 - 4
                End If
            Next j2
        Next i2
    End If
    For i2 = 1 To 16
        If g(i2) <> h(i2) Then fc = True
        If g(i2) = 0 Then fd = True
    Next i2
    If fc = True Or fc = False And fd = False Then Call sj2
    Call sc2
End Sub

Sub sc1()
    For i1 = 0 To 15
        If e(i1 + 1) <> f(i1 + 1) Then
            If e(i1 + 1) = 0 Then
                Image1(i1).Picture = Nothing
            Else
                Image1(i1) = LoadResPicture(100 + e(i1 + 1), vbResBitmap)
            End If
        End If
    Next i1
    Call zhuangtai1
End Sub

Sub sc2()
    For i2 = 0 To 15
        If g(i2 + 1) <> h(i2 + 1) Then
            If g(i2 + 1) = 0 Then
                Image2(i2).Picture = Nothing
            Else
                Image2(i2) = LoadResPicture(100 + g(i2 + 1), vbResBitmap)
            End If
        End If
    Next i2
    Call zhuangtai2
End Sub

Sub zhuangtai1()
    s3 = ""
    DeveloperMode.List1.Clear
    DeveloperMode.List1.AddItem "左侧玩家方块数字监测"
    For i1 = 1 To 16
        s3 = s3 + " e(" + MainForm.zfc(i1) + ")=" + MainForm.zfc(e(i1))
        If i1 Mod 4 = 0 Then
            DeveloperMode.List1.AddItem s3
            s3 = ""
        End If
    Next i1
    DeveloperMode.List1.AddItem "max1=" + Str(max1)
End Sub

Sub zhuangtai2()
    s4 = ""
    DeveloperMode.List2.Clear
    DeveloperMode.List2.AddItem "右侧玩家方块数字监测"
    For i2 = 1 To 16
        s4 = s4 + " g(" + MainForm.zfc(i2) + ")=" + MainForm.zfc(g(i2))
        If i2 Mod 4 = 0 Then
            DeveloperMode.List2.AddItem s4
            s4 = ""
        End If
    Next i2
    DeveloperMode.List2.AddItem "max2=" + Str(max2)
End Sub

Private Sub Command1_Click()
    If Val(Text1.Text) > 32767 Or Val(Text1.Text) < 1 Or Val(Text1.Text) <> Int(Val(Text1.Text)) Then
        MsgBox "请输入一个在1-32767之间的整数", 0 + vbExclamation + vbSystemModal, "参数错误"
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
    End If
    shijian = Val(Text1.Text)
    Call fuwei
    Shape1.Visible = True
    Timer1.Enabled = True
    Command1.Visible = False
    Image1(0).Width = 120
    Image1(0).Height = 120
    Image2(0).Width = 120
    Image2(0).Height = 120
    For jieshu = 0 To 15
        Image1(jieshu).Visible = True
        Image2(jieshu).Visible = True
    Next jieshu
End Sub
Private Sub Form_keypress(keyascii As Integer)
    If f5 = False Then
        Call MainProgress1(keyascii)
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If f5 = False Then
        Call MainProgress2(KeyCode)
    End If
    If KeyCode = 27 Then
        Timer1.Enabled = False
        Unload duorenmoshi
        DeveloperMode.List1.Clear
        DeveloperMode.List2.Clear
    End If
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call Command1_Click
End Sub

Sub sj1()
    Randomize
    fx1 = False: fy1 = False
    k1 = 0
    Erase s1
    For i1 = 1 To 16
        If e(i1) = 0 Then k1 = k1 + 1: s1(k1) = i1
    Next i1
    If k1 <> 0 Then
        t1 = Int(Rnd() * k1) + 1
        e(s1(t1)) = 1
        dh1 = s1(t1)
    End If
    For i1 = 1 To 4
        For j1 = 0 To 2
            If e(j1 * 4 + i1) = e((j1 + 1) * 4 + i1) Then fy1 = True
        Next j1
    Next i1
    For j1 = 0 To 3
        For i1 = 1 To 3
            If e(j1 * 4 + i1) = e(j1 * 4 + i1 + 1) Then fx1 = True
        Next i1
    Next j1
    For i1 = 1 To 16
        If e(i1) > max1 Then max1 = e(i1)
    Next i1
End Sub

Sub sj2()
    Randomize
    fx2 = False: fy2 = False
    k2 = 0
    Erase s2
    For i2 = 1 To 16
        If g(i2) = 0 Then k2 = k2 + 1: s2(k2) = i2
    Next i2
    If k2 <> 0 Then
        t2 = Int(Rnd() * k2) + 1
        g(s2(t2)) = 1
        dh2 = s2(t2)
    End If
    For i2 = 1 To 4
        For j2 = 0 To 2
            If g(j2 * 4 + i2) = g((j2 + 1) * 4 + i2) Then fy2 = True
        Next j2
    Next i2
    For j2 = 0 To 3
        For i2 = 1 To 3
            If g(j2 * 4 + i2) = g(j2 * 4 + i2 + 1) Then fx2 = True
        Next i2
    Next j2
    For i2 = 1 To 16
        If g(i2) > max2 Then max2 = g(i2)
    Next i2
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    duorenmoshi.Show
    Text1.SetFocus
End Sub
Private Sub fuwei()
    Shape1.Top = -100
    Shape1.Visible = False
    Shape1.Left = duorenmoshi.Width / Screen.TwipsPerPixelX / 2 - Shape1.Width / 2
    Shape2.Left = duorenmoshi.Width / Screen.TwipsPerPixelX / 2 - Shape2.Width / 2
    Shape2.Visible = False
    f5 = True
    t3 = 0: t4 = 0: max1 = 0: max2 = 0
    duorenmoshi.Cls
    Timer1.Interval = 1
    Erase e
    Erase f
    Erase g
    Erase h
    Text1.Visible = False
    For ur = 0 To 15
        Image1(ur) = Nothing
        Image2(ur) = Nothing
    Next ur
    Call sj1
    Call sj2
    Call sc1
    Call sc2
End Sub
Private Sub bijiao()
    If max1 > max2 Then
        For jieshu = 0 To 15
            If jieshu > 0 Then Image1(jieshu).Visible = False
            Image2(jieshu).Visible = False
        Next jieshu
        Image1(0).Width = Image1(0).Width * 4
        Image1(0).Height = Image1(0).Height * 4
        Image1(0) = LoadResPicture(98, vbResBitmap)
    ElseIf max1 < max2 Then
        For jieshu = 0 To 15
            If jieshu > 0 Then Image2(jieshu).Visible = False
            Image1(jieshu).Visible = False
        Next jieshu
        Image2(0).Width = Image2(0).Width * 4
        Image2(0).Height = Image2(0).Height * 4
        Image2(0) = LoadResPicture(98, vbResBitmap)
    ElseIf max1 = max2 Then
        MsgBox "最大数相同,平局", 0 + vbSystemModal
    End If
    Text1.SetFocus
End Sub

Private Sub Timer1_Timer()
    If f5 = True And Shape1.Top < 20 Then
        Shape1.Top = Shape1.Top + 1
    ElseIf f5 = True And Shape1.Top >= 20 Then
        t3 = t3 + 1
        If t3 > 2 Then
            Shape2.Left = Shape2.Left - 4
            Shape2.Width = Shape2.Width + 8
            Shape2.BackColor = RGB(255 - t3, 0, t3)
            If Shape2.Left <= 10 Then f5 = False
        End If
    End If
    If Shape1.Top = 20 And t3 = 1 Then
        Shape2.Top = Shape1.Top + 35
        Shape2.Visible = True
    End If
    If t4 = 1 Then Timer1.Interval = shijian: Print "开始时间": Print Now
    If Shape2.Width > Shape1.Width \ 2 And f5 = False Then
        Shape2.Left = Shape2.Left + 1
        Shape2.Width = Shape2.Width - 2
        Shape2.BackColor = RGB(0.375 * t4, 0, (255 - 0.375 * t4))
        t4 = t4 + 1
        If Shape2.Width <= Shape1.Width \ 2 Then f5 = True: Timer1.Enabled = False: Command1.Visible = True: Text1.Visible = True: Print "结束时间": Print Now:  Call bijiao
    End If
    If max1 <> 0 And max2 <> 0 And fy1 = False And fy2 = False And fx1 = False And fx2 = False Then f5 = True: Timer1.Enabled = False: Command1.Visible = True: Text1.Visible = True: Print "结束时间": Print Now: Call bijiao
End Sub
