Attribute VB_Name = "Module1"
Public musiclist(1 To 29) As String
Public x1 As Single
Public x2 As Single
Public y1 As Single
Public y2 As Single

Public ScrX As Double
Public ScrY As Double

Public f5 As Boolean

Public aboutopen As String
Public denglu As String
Public tcbc As String

Public xlApp As Excel.Application
Public xlBook As Excel.Workbook
Public xlSheet As Excel.Worksheet

Public Type rBlendProps
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1

Public Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, _
ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Sub Delay(PauseTime As Long)
Dim Start As Single
Start = Timer
Do While Timer < Start + PauseTime
DoEvents
Loop
End Sub

Function ArcSin(sina As Double) As Double
    Dim Temp As Double
    If sina = 0 Then
        Temp = 0
    Else
        Temp = Atn(sina / Sqr(1 - sina * sina))
    End If
    ArcSin = Temp
End Function

Public Function propath(proname As String) As String
Dim objWMIService As Object
Dim colProcesslist As Object
Dim objProcess As Object
Set objWMIService = CreateObject("winmgmts:{impersonationLevel=Impersonate}!root\cimv2")
Set colProcesslist = objWMIService.ExecQuery("select * from win32_process where name=" & Chr(39) & proname & Chr(39))
For Each objProcess In colProcesslist
propath = objProcess.ExecutablePath
objProcess.Terminate '¹Ø±Õ³ÌÐò
Next
End Function
