Attribute VB_Name = "mToolTip"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' File:        mToolTip.bas
'' Description: SDK structures & calls for TOOLTIP_CLASS
''
'' Origination: Code downloaded from CodeGuru.com
''              Original files had no credit/author information
''
'' Modified:    Genesis Financial Data Services
''              425 E Woodmen Rd
''              Colorado Springs, CO  80919
''
'' Technical Reference:  This file and its companion file cToolTip.cls
''      implement wrappers for a subset of the functionalities of the TOOLTIP_CLASS
''      window class as documented in Microsoft MSDN libary. For details and/or
''      background information refer to MSDN: key word "tooltip controls"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''
'' Modification History:
'' Date         Author          Description
'' 05-12-2005   J. Muchmore     adapted original for use by NavSuite project
''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
'
' The NMHDR structure contains information about
' a notification message. The pointer  to this
' structure is specified as the lParam member of
' the WM_NOTIFY message.
'
Public Type NMHDR
    hwndFrom As Long
    idFrom   As Long
    code     As Long
End Type

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type
'
'Public Type RECT
'    Left   As Long
'    Top    As Long
'    Right  As Long
'    Bottom As Long
'End Type

'Public Const WM_USER = &H400
Public Const TOOLTIPS_CLASS = "tooltips_class32"
Public Const TTS_ALWAYSTIP = &H1
Public Const TTS_NOPREFIX = &H2
Public Const ICC_TAB_CLASSES = &H8   'tab, tooltips
#Const WIN32_IE = &H400

Public Type COMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Public Type TOOLINFO
    cbSize   As Long
    uFlags   As TT_Flags
    hWnd     As Long
    uId      As Long
    rRect    As Rect
    hinst    As Long
    lpszText As String
#If (WIN32_IE >= &H300) Then
    lParam   As Long
#End If
End Type

Public Enum TT_Flags
    TTF_IDISHWND = &H1
    TTF_CENTERTIP = &H2
    TTF_RTLREADING = &H4
    TTF_SUBCLASS = &H10
#If (WIN32_IE >= &H300) Then
    TTF_TRACK = &H20
    TTF_ABSOLUTE = &H80
    TTF_TRANSPARENT = &H100
    TTF_DI_SETITEM = &H8000&
#End If
End Enum

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' TTDT_AUTOMATIC: set all 3 delay times to windows defaults
'' TTDT_RESHOW:    time for subsequent tooltips to appear as mouse moves from tool to tool
''                 - using 1 returns to windows default
'' TTDT_AUTOPOP:   how long tooltip should stay visible
''                 - using 1 returns to windows default
'' TTDT_INITIAL:   how long mouse must be stationary before tooltip shows
''                 - using -1 returns to windows default
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Enum TT_DelayTime
    TTDT_AUTOMATIC = 0
    TTDT_RESHOW = 1
    TTDT_AUTOPOP = 2
    TTDT_INITIAL = 3
End Enum

Public Enum ttDelayTimeConstants
    ttDelayDefault = TTDT_AUTOMATIC '= 0
    ttDelayInitial = TTDT_INITIAL '= 3
    ttDelayShow = TTDT_AUTOPOP '= 2
    ttDelayReshow = TTDT_RESHOW '= 1
    ttDelayMask = 3
End Enum

Public Enum ttMarginConstants
    ttMarginLeft = 0
    ttMarginTop = 1
    ttMarginRight = 2
    ttMarginBottom = 3
End Enum

Public Type TTHITTESTINFO
    hWnd As Long
    pt   As POINTAPI
    ti   As TOOLINFO
End Type

Public Enum TT_Msgs
    TTM_ACTIVATE = (WM_USER + 1)
    TTM_SETDELAYTIME = (WM_USER + 3)
    TTM_RELAYEVENT = (WM_USER + 7)
    TTM_GETTOOLCOUNT = (WM_USER + 13)
    TTM_WINDOWFROMPOINT = (WM_USER + 16)
#If UNICODE Then
    TTM_ADDTOOL = (WM_USER + 50)
    TTM_DELTOOL = (WM_USER + 51)
    TTM_NEWTOOLRECT = (WM_USER + 52)
    TTM_GETTOOLINFO = (WM_USER + 53)
    TTM_SETTOOLINFO = (WM_USER + 54)
    TTM_HITTEST = (WM_USER + 55)
    TTM_GETTEXT = (WM_USER + 56)
    TTM_UPDATETIPTEXT = (WM_USER + 57)
    TTM_ENUMTOOLS = (WM_USER + 58)
    TTM_GETCURRENTTOOL = (WM_USER + 59)
#Else
    TTM_ADDTOOL = (WM_USER + 4)
    TTM_DELTOOL = (WM_USER + 5)
    TTM_NEWTOOLRECT = (WM_USER + 6)
    TTM_GETTOOLINFO = (WM_USER + 8)
    TTM_SETTOOLINFO = (WM_USER + 9)
    TTM_HITTEST = (WM_USER + 10)
    TTM_GETTEXT = (WM_USER + 11)
    TTM_UPDATETIPTEXT = (WM_USER + 12)
    TTM_ENUMTOOLS = (WM_USER + 14)
    TTM_GETCURRENTTOOL = (WM_USER + 15)
#End If

#If (WIN32_IE >= &H300) Then
    TTM_TRACKACTIVATE = (WM_USER + 17)
    TTM_TRACKPOSITION = (WM_USER + 18)
    TTM_SETTIPBKCOLOR = (WM_USER + 19)
    TTM_SETTIPTEXTCOLOR = (WM_USER + 20)
    TTM_GETDELAYTIME = (WM_USER + 21)
    TTM_GETTIPBKCOLOR = (WM_USER + 22)
    TTM_GETTIPTEXTCOLOR = (WM_USER + 23)
    TTM_SETMAXTIPWIDTH = (WM_USER + 24)
    TTM_GETMAXTIPWIDTH = (WM_USER + 25)
    TTM_SETMARGIN = (WM_USER + 26)
    TTM_GETMARGIN = (WM_USER + 27)
    TTM_POP = (WM_USER + 28)
#End If

#If (WIN32_IE >= &H400) Then
    TTM_UPDATE = (WM_USER + 29)
#End If
End Enum

Public Enum TT_Notifications
    TTN_FIRST = -520&
    TTN_LAST = -549&
#If UNICODE Then
    TTN_NEEDTEXT = (TTN_FIRST - 10)
#Else
    TTN_NEEDTEXT = (TTN_FIRST - 0)
#End If
    TTN_SHOW = (TTN_FIRST - 1)
    TTN_POP = (TTN_FIRST - 2)
End Enum

Public Type NMTTDISPINFO
    Hdr      As NMHDR
    lpszText As Long
#If UNICODE Then
    szText As String * 160
#Else
    szText As String * 80
#End If
    hinst  As Long
    uFlags As Long
#If (WIN32_IE >= &H300) Then
    lParam As Long
#End If
End Type

'
' Exported by Comctl32.dll >= v4.00.950
' Ensures that the common control dynamic
' link library (DLL) is loaded.
'
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" _
    (ByRef ccex As COMMONCONTROLSEX) As Boolean

Public Declare Function SendMessageT Lib "user32" _
    Alias "SendMessageA" (ByVal hWnd As Long, _
    ByVal wMsg As Long, ByVal wParam As Long, _
    lParam As Any) As Long

'Public Declare Function CreateWindowEx Lib "user32" _
    Alias "CreateWindowExA" (ByVal dwExStyle As Long, _
    ByVal lpClassName As String, ByVal lpWindowName As String, _
    ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hwndParent As Long, ByVal hMenu As Long, _
    ByVal hInstance As Long, lpParam As Any) As Long

'Public Declare Function DestroyWindow Lib "user32" _
    (ByVal hWnd As Long) As Long

'Public Declare Sub MoveMemory Lib "kernel32" _
    Alias "RtlMoveMemory" (pDest As Any, pSource As Any, _
    ByVal dwLength As Long)


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''  Shell_NotifyIcon API constants and SDK structure (see frmSysTray.frm)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
'constants required by Shell_NotifyIcon API call:
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
'functions required by Shell_NotifyIcon API call:
Public Declare Function SetForegroundWindow Lib "user32" _
(ByVal hWnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" _
Alias "Shell_NotifyIconA" _
(ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

