VERSION 5.00
Begin VB.Form UpdateForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������־"
   ClientHeight    =   5385
   ClientLeft      =   30
   ClientTop       =   360
   ClientWidth     =   8475
   Icon            =   "Form7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form7.frx":038A
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.TextBox Text1 
      Height          =   5400
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form7.frx":0714
      Top             =   0
      Width           =   8500
   End
End
Attribute VB_Name = "UpdateForm"
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
    Text1.Text = "�汾:V6.1.0��ʽ�� 2022.7.6" + vbCrLf + "�޸�:" + vbCrLf + "�޸��ڿ�����ѡ�������˶�սʱList1���г���5*5����Ԫ�ؼ�������" + vbCrLf + "�޸����˶�սģʽʤ���ж����ƣ����ڲ�����Ҫռ�������Ҳ����ƶ�������ȷ�ж�" + vbCrLf + vbCrLf + "����:" + vbCrLf + "�޸Ŀ�����ѡ��� ��ʾ/�������ֿؼ� ����ʾ����Ϊ��������仯" + vbCrLf + "�������˶�սģʽ���������ڿ�����ѡ�������List����" + vbCrLf + "����������ѡ�������List����ģʽ�仯�Զ�������С" + vbCrLf + "����Ϊ��ģʽ����ر�ʱ�Կ�����ѡ��List�����գ���ȫ����Ч��" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V6.0.0��ʽ�� 2022.1.31" + vbCrLf + "�޸�:" + vbCrLf + "�޸��ڿ����������غ�ʹ�����²��ƶ����ڿ�����ѡ��ʹ�ò���һ������������⹦��ʱ����ĳһ��������ʾ�Ŵ󶯻�������" + vbCrLf + "�޸��ڿ����������غ��ڿ�����ѡ������ظ�ʹ�ò���һ������������⹦��ʱ����ķŴ󶯻���ǰ��ֹ���·����С���޷��ָ�������" + vbCrLf + "�޸����ؽ���������100%ʱ������δ����߿������" + vbCrLf + vbCrLf + "����:" + vbCrLf + "��ȫ�˷��黬��������ʹ�÷����ƶ�����˳��������" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V5.4.4��ʽ�� 2022.1.24" + vbCrLf + "�޸�:" + vbCrLf + "�޸���6*6�еĳ�����ʾ���������" + vbCrLf + "�޸��ڸ��ֱַ����¼��ض���LOGO��ʾƫ�ƻ�ȫ������" + vbCrLf + "�޸���������Ϸ�м��ش浵��δ���в���ʱ��һ������һ����ť���õ�����" + vbCrLf + "�޸�����ģʽ�������ƶ���������˸������" + vbCrLf + "�޸�����������Ϣ����" + vbCrLf + "�޸ķ�ȫ��ģʽ�˳��߼�:����ESC�˳�ʱ�����˳�֮��ѡ�����˳�֮������ʾ��" + vbCrLf + "�޸ķ�ȫ��ģʽ�˳��߼�:����������������Ͻ��˳���ťʱ�����˳�֮������ʾ��" _
+ vbCrLf + "���Ŀ�����ѡ��ʵʱ������Ŀ" + vbCrLf + "�޸�����������Ϳ�����ѡ�����Ĳ�͸���ȵ�����ΧΪ[20%,100%]��[60%,100%]" + vbCrLf + vbCrLf + "����:" + vbCrLf + "�����˷�ȫ��ģʽ" + vbCrLf + "�ٴ��Ż������߼��������Ϸ��Ӧ�ٶ�" + vbCrLf + "��ԭ�л��������Ĵ浵����ʱ�����ݼ�顢�ݴ���ͱ�����ʾ" + vbCrLf + "��������ģʽʱ��δ����֮ǰ˫�����޷��ƶ����������ǰ������Ϸ" + vbCrLf + "��������Ϸ����ʱ����Ϸ˵�����壬����ѡ���´��Ƿ���ʾ��ԭ�Ҽ��˵����뷽ʽ����" _
+ vbCrLf + "����10������MineCraft1.18��BGM" + vbCrLf + "�·���һ����˳���и�����Ϸ�Ҽ��˵�" + vbCrLf + "�޸��˲���������Ϣ" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V5.4��ʽ�� 2021.1.9" + vbCrLf + "�޸�:" + vbCrLf + "�����޸������ϵͳ�ֱ������õ��¼�����ϷʱLogo��ʾ����ȫ������" + vbCrLf + "�޸���ȡ�浵�������Ϸģʽ��ѡ����Ӧ״̬����ȷ������" + vbCrLf + "�޸ĵ���Ϸ�����ϴ�������ʱ������ѡ�����" + vbCrLf + "�޸��˳�ѡ����һ:����4*4��6*6����Escʱ���ᵯ�� �˳�֮һ ѯ���Ƿ��˳���Ϸ" + vbCrLf + "�޸��˳�ѡ�����:�������������水��Escʱ�����ᵯ�� �˳�֮�� �����˳���Ϸ" _
+ vbCrLf + "�޸ı���浵��ʵʱ�����ʱ��Ϊ���ƣ���ͬ·���µ�saves�ļ���Ϊ����λ��" + vbCrLf + "�޸Ķ�ȡ�浵Ĭ��·��Ϊsaves�ļ���" + vbCrLf + "�޸�����������Ľ��뷽ʽ�����ڿ���ͨ���س���������Ϸ" + vbCrLf + "�޸Ŀ�����ѡ����������½��͸���ȵ��ڣ����Ĳ������ݺ͹���" + vbCrLf + "�޸Ĳ˵�����,����Ϸ������������ݺͲ���,��������հ״����������ں͸���" + vbCrLf + "�޸Ľ��沼�ֺ�UI�����ִ����Ż�" + vbCrLf + "�޸Ĳ���������Ϣ" + vbCrLf + vbCrLf + "����:" + vbCrLf + "������˶�սģʽ" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V5.2��ʽ�� 2020.7.19" + vbCrLf + "�޸�:" + vbCrLf + "�޸�����ʾ���ֱ��ʻ��ݺ�Ȳ�ͬ���µĽ������" + vbCrLf + "�޸�������������û����Ϸͼ�������" + vbCrLf + "�����Ѷ�ѡ����ģʽѡ��˵�״̬��ȡ������ʧ�ܵķ��أ�����ͨ��������ǻ�ֱ���Ϸģʽ" + vbCrLf + vbCrLf + "����:" + vbCrLf + "����15������MineCraft��BGM��ȡ���򿪴����Զ��и�" + vbCrLf + "����ģʽ�İ�" + vbCrLf + "�����������أ���ֹ��" + vbCrLf + "�Ż����ؽ���" + vbCrLf + "�����Ĵ����Ż�" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V5.1��ʽ�� 2020.7.10" + vbCrLf + "�޸�:" + vbCrLf + "������������ر���Ϸʱǿ�ƹر����зֽ���" + vbCrLf + "������Ϸ״̬��ʾ" + vbCrLf + vbCrLf + "����:" + vbCrLf + "������ʷ��¼���ܣ�ͨ������һ��������" + vbCrLf + "excel�򿪷��ڼ��ؽ��棬���ڵļ����������" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V4.5��ʽ�� 2020.7.9" + vbCrLf + "�޸�:" + vbCrLf + "�޸�3.3�汾������ѡ�ťЧ�����ô���" + vbCrLf + vbCrLf + "����:" + vbCrLf + "����������/�������������ڼ�����Ϸ���ں󣬰������������ƶ��������ɿ���������Զ�����������겢�ƶ�����" + vbCrLf + "�����浵�ĵ���͵���" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V3.3��ʽ�� 2020.7.8" + vbCrLf + "����:" + vbCrLf + "����4������MineCraft��BGM����ˣ�����л�ѭ������" + vbCrLf + "����5*5,6*6�淨������ര��ʱ��" + vbCrLf + "��������ģʽ����ģʽ�µ�ÿһ��������1%�ĸ��������ֳ��ֲ���ռ��һ���ո��������޷��ϳɣ�������������һ����ˣ" + vbCrLf + "�Ż�����ṹ����Ӧ�����������" + vbCrLf + _
"ʹ����Դ�⣬�����ⲿ�����ļ�" + vbCrLf + "ȡ��ListBox�棬ȡ������ͼƬģʽ" + vbCrLf + "����Logo���������ؽ��棨��������㲻���������Ϸ�ᷢ��Logo���й��ɵĺ�����" + vbCrLf + "��ȫ������ƵĽ��沼��" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V2.2��ʽ�� 2020.6.28" + vbCrLf + "�޸�:" + vbCrLf + "�޸�����λ��������λʱ�Ű淢������" + vbCrLf + vbCrLf + _
"����:" + vbCrLf + "��ListBox��Ļ�������������ͼƬģʽ���ֻ�ͼƬģʽ" + vbCrLf + "ListBox�潫����0��Ϊ�ո���ʾ" + vbCrLf + "������ҷָ��" + vbCrLf + "ȡ��TextBox�����룬���ڿ�����Form��ֱ�Ӱ�����" + vbCrLf + "���Ӵ�СдWwAaSsDd����" + vbCrLf + "���ӿ�����ѡ�Ѹ�����ü������" + vbCrLf + "�����Ѷȼ���" + vbCrLf + "���Ӳ�����ʾ" + vbCrLf + "�����Ż�" _
+ vbCrLf + vbCrLf + vbCrLf + _
"�汾:V1.0��ʽ�� 2020.6.21" + vbCrLf + "�����汾"
End Sub



