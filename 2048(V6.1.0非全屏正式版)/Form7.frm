VERSION 5.00
Begin VB.Form UpdateForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "更新日志"
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
   StartUpPosition =   2  '屏幕中心
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
    Text1.Text = "版本:V6.1.0正式版 2022.7.6" + vbCrLf + "修复:" + vbCrLf + "修复在开发者选项点击多人对战时List1框中出现5*5数组元素监测的问题" + vbCrLf + "修复多人对战模式胜负判定机制，现在不再需要占满格子且不能移动就能正确判定" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "修改开发者选项的 显示/隐藏音乐控件 的显示文字为随操作而变化" + vbCrLf + "新增多人对战模式的数组监测于开发者选项的两个List框中" + vbCrLf + "新增开发者选项的两个List框随模式变化自动调整大小" + vbCrLf + "新增为各模式窗体关闭时对开发者选项List框的清空（非全屏无效）" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V6.0.0正式版 2022.1.31" + vbCrLf + "修复:" + vbCrLf + "修复在开启滑动开关后使用上下步移动或在开发者选项使用产生一致数和玩死检测功能时存在某一个方块显示放大动画的问题" + vbCrLf + "修复在开启滑动开关后在开发者选项快速重复使用产生一致数和玩死检测功能时方块的放大动画提前终止导致方块变小且无法恢复的问题" + vbCrLf + "修复加载进度条到达100%时进度条未到达边框的问题" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "补全了方块滑动动画，使得方块移动更加顺滑且清晰" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V5.4.4正式版 2022.1.24" + vbCrLf + "修复:" + vbCrLf + "修复在6*6中的出现显示错误的问题" + vbCrLf + "修复在各种分辨率下加载动画LOGO显示偏移或不全的问题" + vbCrLf + "修复在任意游戏中加载存档后未进行操作时上一步和下一步按钮可用的问题" + vbCrLf + "修复多人模式下数字移动卡顿且闪烁的问题" + vbCrLf + "修复部分文字信息错误" + vbCrLf + "修改非全屏模式退出逻辑:按下ESC退出时弹出退出之二选择框和退出之三的提示框" + vbCrLf + "修改非全屏模式退出逻辑:点击主交互界面右上角退出按钮时弹出退出之三的提示框" _
+ vbCrLf + "增改开发者选项实时检测的项目" + vbCrLf + "修改主交互界面和开发者选项界面的不透明度调整范围为[20%,100%]和[60%,100%]" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "补充了非全面模式" + vbCrLf + "再次优化代码逻辑，提高游戏响应速度" + vbCrLf + "在原有基础上增改存档导入时的数据检查、容错处理和报错提示" + vbCrLf + "新增多人模式时间未结束之前双方都无法移动的情况下提前结束游戏" + vbCrLf + "新增在游戏启动时打开游戏说明窗体，可以选择下次是否显示，原右键菜单进入方式不变" _
+ vbCrLf + "新增10首来自MineCraft1.18的BGM" + vbCrLf + "下放下一步和顺序切歌至游戏右键菜单" + vbCrLf + "修改了部分文字信息" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V5.4正式版 2021.1.9" + vbCrLf + "修复:" + vbCrLf + "部分修复因操作系统分辨率设置导致加载游戏时Logo显示不完全的问题" + vbCrLf + "修复读取存档后更改游戏模式的选项响应状态不正确的问题" + vbCrLf + "修改当游戏界面上存在游乐时，保存选项不可用" + vbCrLf + "修改退出选项其一:当在4*4或6*6按下Esc时将会弹出 退出之一 询问是否退出游戏" + vbCrLf + "修改退出选项其二:当在主交互界面按下Esc时，将会弹出 退出之二 真正退出游戏" _
+ vbCrLf + "修改保存存档以实时计算机时间为名称，以同路径下的saves文件夹为保存位置" + vbCrLf + "修改读取存档默认路径为saves文件夹" + vbCrLf + "修改主交互界面的进入方式，现在可以通过回车键进入游戏" + vbCrLf + "修改开发者选项，新增密码登陆和透明度调节，增改部分内容和功能" + vbCrLf + "修改菜单部署,在游戏界面分立出内容和操作,在主界面空白处分立出关于和更新" + vbCrLf + "修改界面布局和UI，部分代码优化" + vbCrLf + "修改部分文字信息" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "加入多人对战模式" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V5.2正式版 2020.7.19" + vbCrLf + "修复:" + vbCrLf + "修复因显示器分辨率或纵横比不同导致的界面混乱" + vbCrLf + "修复了桌面任务栏没有游戏图标的问题" + vbCrLf + "调整难度选项与模式选项菜单状态，取消更改失败的返回，可以通过窗体活动与非活动分辨游戏模式" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "新增15首来自MineCraft的BGM，取消打开窗体自动切歌" + vbCrLf + "游乐模式改版" + vbCrLf + "新增滑动开关，防止误触" + vbCrLf + "优化加载界面" + vbCrLf + "大量的代码优化" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V5.1正式版 2020.7.10" + vbCrLf + "修复:" + vbCrLf + "在主交互界面关闭游戏时强制关闭所有分界面" + vbCrLf + "调整游戏状态显示" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "新增历史记录功能，通过”上一步”倒退" + vbCrLf + "excel打开放在加载界面，现在的加载是真的了" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V4.5正式版 2020.7.9" + vbCrLf + "修复:" + vbCrLf + "修复3.3版本开发者选项按钮效果设置错误" + vbCrLf + vbCrLf + "更新:" + vbCrLf + "新增触摸屏/鼠标操作方法：在激活游戏窗口后，按下鼠标左键后移动鼠标随后松开，程序会自动计算横纵坐标并移动方块" + vbCrLf + "新增存档的导入和导出" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V3.3正式版 2020.7.8" + vbCrLf + "更新:" + vbCrLf + "新增4首来自MineCraft的BGM，玩耍过程中会循环播放" + vbCrLf + "新增5*5,6*6玩法，步入多窗体时代" + vbCrLf + "新增游乐模式，该模式下的每一步都会有1%的概率让游乐出现并且占据一个空格以至于无法合成，让游乐陪我们一起玩耍" + vbCrLf + "优化代码结构，适应更大的运算量" + vbCrLf + _
"使用资源库，减少外部配置文件" + vbCrLf + "取消ListBox版，取消人像图片模式" + vbCrLf + "新增Logo，新增加载界面（雾）（如果你不点击进入游戏会发现Logo在有规律的呼吸）" + vbCrLf + "完全重新设计的界面布局" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V2.2正式版 2020.6.28" + vbCrLf + "修复:" + vbCrLf + "修复数字位数大于四位时排版发生错误" + vbCrLf + vbCrLf + _
"更新:" + vbCrLf + "在ListBox版的基础上增加人像图片模式和手绘图片模式" + vbCrLf + "ListBox版将数字0改为空格显示" + vbCrLf + "添加左右分割框" + vbCrLf + "取消TextBox框输入，现在可以在Form中直接按键盘" + vbCrLf + "增加大小写WwAaSsDd输入" + vbCrLf + "增加开发者选项，迅速设置极端情况" + vbCrLf + "增加难度级数" + vbCrLf + "增加步数显示" + vbCrLf + "代码优化" _
+ vbCrLf + vbCrLf + vbCrLf + _
"版本:V1.0正式版 2020.6.21" + vbCrLf + "初代版本"
End Sub



