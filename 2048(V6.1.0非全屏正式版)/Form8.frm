VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form about 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��Ϸ˵��"
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
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5640
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�رգ��´���Ȼ��ʾ"
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   7560
      Width           =   5895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "�رգ��´β�����ʾ"
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
    Text1.Text = "2048�������򣺿������з�����ͬһ�����ƶ����ڸ÷��������������Ե��������ڻ���յ���ͬ���ַ����ϲ���Ϊ���ǵĺͣ���û������򱣳ֲ��䣬ȫ���ϲ���ɺ��ظ÷�����ף�������и������������޷�����Ժϲ����ж���Ϸ����" _
+ vbCrLf + vbCrLf + "����Ϸ�����Ҽ����Ժ�����Ϸ�˵����������²��ƶ��������ȡ�浵����������Ϸ�ļ���\saves\��������ʱ����������txt�ı��ļ���������/�رջ������ء�������Ϸģʽ�����䡢���֣���������Ϸ�Ѷȣ���ͨ�����ѡ��������Լ�˳���и裨��29�ף�" + vbCrLf + vbCrLf + "�����������Ҽ����Ժ�����Ϣ�˵���������Ϸ˵���������壩�Լ�2048������־" _
+ vbCrLf + vbCrLf + "ֻҪ��������Ϸ�ѶȻ�����Ϸģʽ�������¿�ʼ���������Ѷȼ����ϵ��Ѷ������ɵ����ֻ��벽���������ʹ��Ϸ������һ����ѧ����Ϣ" + vbCrLf + Chr$(10) + "���°�����ģʽ��������ģʽ�����ж���ÿһ�����и��ʣ�Ĭ��Ϊ5%�������ĳһ�����ӳ������֣����һ��������ʱ����������������ͬ�������С���ָ���Ҵ���ϣ�����߾���" + vbCrLf + "���֣���Զ���񣡣�0~0��" _
+ vbCrLf + vbCrLf + "������Ϸ����:" + vbCrLf + "˫��ͼ��4��6���벻ͬ��С��Ϸ������Ϸ���������޷��ƶ�����Ϸ����" _
+ vbCrLf + vbCrLf + "������Ϸ����:" + vbCrLf + "˫�������־������˶�ս�����ı������������ֺ�س����·���ť�Կ�ʼ�����ֲ���ʱ�������ȵ��Ϸ���������ɴ���ɫ���ʱ����˫����ť�������ִ�Сд��WASD���������ҷ���������ж�ս���������������ʧ������Բ��˫�����޷��ƶ�����ʱ��Ϸ��������˫���ϳɵ����ֵ��Ϊ�Ƚ�ʤ���ı�׼�������Ͻǻ��ӡ�����ζ�ս��ʼĩ��Ϣ" _
+ vbCrLf + vbCrLf + "�˳���Ϸ����:" + vbCrLf + "����Ϸ�����ϰ�ESC�ᵯ���Ƿ񱣴�浵����/����ر���Ϸ���壬���������ϰ�ESC�ᵯ���Ƿ��˳���Ϸ��ȷ�Ϻ���ʽ�˳�" _
+ vbCrLf + vbCrLf + "��������:" + vbCrLf + "1-����Ϸ�����Ҽ���������أ�ֱ��ʹ����갴ס-����-�ɿ�����ʹ�ô������㻮" + vbCrLf + "2-����Ҫ��ˣ�Ĵ��壬�������̣������ִ�Сд�����������������뷨��" + vbCrLf + "  W��������" + vbCrLf + "  A��������" + vbCrLf + "  S��������" + vbCrLf + "  D��������" _
+ vbCrLf + vbCrLf + "��ǰ�汾:2048V6.0.0��ʽ��" + vbCrLf + "��Ϸ����:�ѿ�֮��&����֮��" + vbCrLf + "�����ȫ��Ȩ�����ѿ�֮��&����֮������" + vbCrLf + "δ���������ɣ���ֹ�Ա���������κ���ʽ�ķ������޸ġ������롢�ƽ����Ϊ" + vbCrLf + "����������ճ����֣���ֹ������������κ���ʽ����ҵ��Ϊ"
End Sub
