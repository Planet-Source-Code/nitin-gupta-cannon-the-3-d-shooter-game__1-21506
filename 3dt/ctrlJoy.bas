Attribute VB_Name = "ctrlJoy"
Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Public Const MAXPNAMELEN = 32  '  max product name length (including NULL)
Public Const JOY_BUTTON1 = &H1
Public Const JOY_BUTTON2 = &H2
Public Const JOYERR_NOERROR = (0)  '  no error
Public Const JOYSTICK1 = 0
Public Const JOYERR_BASE = 160
Public Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7) '  joystick is unplugged
Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)  '  no device driver present
Public Type JOYCAPS
        wMid As Integer
        wPid As Integer
        szPname As String * MAXPNAMELEN
        wXmin As Long
        wXmax As Long
        wYmin As Long
        wYmax As Long
        wZmin As Long
        wZmax As Long
        wNumButtons As Long
        wPeriodMin As Long
        wPeriodMax As Long
End Type
Public Type JOYINFO
        wXpos As Long
        wYpos As Long
        wZpos As Long
        wButtons As Long
End Type



