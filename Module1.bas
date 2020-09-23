Attribute VB_Name = "Module1"
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const HTLEFT = 10
Public Const HTRIGHT = 11
Public Const HTTOP = 12
Public Const HTTOPLEFT = 13
Public Const HTTOPRIGHT = 14
Public Const HTBOTTOM = 15
Public Const HTBOTTOMLEFT = 16
Public Const HTBOTTOMRIGHT = 17

Public Type POINTAPI
    curX As Long
    curY As Long
End Type


Public Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long

Public Declare Function SendMessage Lib "User32" Alias _
                "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                ByVal wParam As Long, lParam As Any) As Long
                
Public Declare Function ReleaseCapture Lib "User32" () As Long

